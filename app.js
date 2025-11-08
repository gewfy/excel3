import * as THREE from "three";
import { AnaglyphEffect } from "three/examples/jsm/effects/AnaglyphEffect.js";

// Constants
const CELL_WIDTH = 100; // Width (X axis) - wide like Excel
const CELL_HEIGHT = 30; // Height (Y axis) - short like Excel rows
const CELL_DEPTH = 100; // Depth (Z axis) - for stacking layers
const GRID_SIZE_Z = 5; // Number of depth layers
const LABEL_OFFSET_X = 90; // Space for row labels on the left
const LABEL_OFFSET_Y = 50; // Space for column labels on top

// Dynamic grid size based on window
const TOOLBAR_HEIGHT = 48; // Height of the toolbar
let GRID_SIZE_X = Math.floor((window.innerWidth - LABEL_OFFSET_X) / CELL_WIDTH); // Columns based on viewport width
let GRID_SIZE_Y = Math.floor(
  (window.innerHeight - TOOLBAR_HEIGHT - LABEL_OFFSET_Y) / CELL_HEIGHT
); // Rows based on viewport height (minus toolbar)

// Application state
const cellData = {}; // Store cell values: "x,y,z" -> value
const cellBackgroundColors = {}; // Store cell background colors: "x,y,z" -> color
const cellTextColors = {}; // Store cell text colors: "x,y,z" -> color
const cellTextBold = {}; // Store bold state: "x,y,z" -> boolean
const cellTextItalic = {}; // Store italic state: "x,y,z" -> boolean
const cellTextStrikethrough = {}; // Store strikethrough state: "x,y,z" -> boolean
const cellFontFamily = {}; // Store font family: "x,y,z" -> font name
let scene, camera, renderer;
let cellMeshes = [];
let textSprites = [];
let labelSprites = [];
let selectedCell = null;
let selectedCellMesh = null;
let selectionOutline = null;
let selectionStart = null; // Starting cell for drag selection
let selectionEnd = null; // Ending cell for drag selection
let isDragging = false;
let actuallyDragged = false; // Track if mouse actually moved during drag
let justFinishedDragging = false;
let isRotating = false;
let isEditingCell = false;
let editingCellCoords = null;
let editingText = "";
let previousMousePosition = { x: 0, y: 0 };
let pivot; // Group to hold all cells for rotation

// Zoom state
let initialPinchDistance = 0;
let lastTouches = [];

// Pan state
let isPanning = false;
let previousPanPosition = { x: 0, y: 0 };

// Anaglyph 3D state
let isAnaglyphMode = false;
let anaglyphEffect = null;
let perspectiveCamera = null;
let orthographicCamera = null; // Store reference to original camera

// Initialize the application
function init() {
  // Create scene
  scene = new THREE.Scene();
  scene.background = new THREE.Color(0xffffff);

  // Create orthographic camera for top-left origin
  const gridWidth = GRID_SIZE_X * CELL_WIDTH;
  const gridHeight = GRID_SIZE_Y * CELL_HEIGHT;

  // Set up camera to view grid from top-left (0,0)
  // Create orthographic camera (default)
  orthographicCamera = new THREE.OrthographicCamera(
    -gridWidth / 2, // Left
    gridWidth / 2, // Right
    gridHeight / 2, // Top
    -gridHeight / 2, // Bottom
    1,
    10000
  );
  // Position camera at grid center looking forward
  const centerX = gridWidth / 2;
  const centerY = -gridHeight / 2;
  orthographicCamera.position.set(centerX, centerY, 3000);
  orthographicCamera.lookAt(centerX, centerY, 0);

  // Create perspective camera (for anaglyph mode)
  perspectiveCamera = new THREE.PerspectiveCamera(
    45,
    window.innerWidth / (window.innerHeight - TOOLBAR_HEIGHT),
    1,
    10000
  );
  perspectiveCamera.position.set(centerX, centerY, 3000);
  perspectiveCamera.lookAt(centerX, centerY, 0);

  // Use orthographic camera by default
  camera = orthographicCamera;

  // Create renderer
  renderer = new THREE.WebGLRenderer({
    antialias: true,
    alpha: true,
    logarithmicDepthBuffer: true, // Better depth precision for transparency
  });
  renderer.setSize(window.innerWidth, window.innerHeight - TOOLBAR_HEIGHT);
  renderer.setPixelRatio(window.devicePixelRatio);
  renderer.sortObjects = true; // Enable sorting for transparent objects
  document.getElementById("canvas-container").appendChild(renderer.domElement);

  // Create anaglyph effect for 3D glasses mode (red-cyan)
  anaglyphEffect = new AnaglyphEffect(renderer);
  anaglyphEffect.setSize(
    window.innerWidth,
    window.innerHeight - TOOLBAR_HEIGHT
  );

  // Create pivot group for rotation
  pivot = new THREE.Group();
  pivot.position.set(0, 0, 0); // Pivot at origin
  pivot.rotation.set(0, 0, 0); // Ensure no initial rotation (flat 2D view)
  scene.add(pivot);

  // Add lighting for 3D shading with vibrant colors
  // Moderate ambient light for base brightness
  const ambientLight = new THREE.AmbientLight(0xffffff, 0.5);
  scene.add(ambientLight);

  // Strong front light (from camera direction) to brighten facing surfaces
  const frontLight = new THREE.DirectionalLight(0xffffff, 1.0);
  frontLight.position.set(0, 0, 1); // From the front (camera direction)
  scene.add(frontLight);

  // Stronger side light for pronounced shading
  const sideLight = new THREE.DirectionalLight(0xffffff, 0.6);
  sideLight.position.set(1, 0.5, 0.3);
  scene.add(sideLight);

  // Top light for depth
  const topLight = new THREE.DirectionalLight(0xffffff, 0.3);
  topLight.position.set(0, 1, 0.5);
  scene.add(topLight);

  // Create grid of cells
  createCellGrid();

  // Create row and column labels
  createLabels();

  // Event listeners
  window.addEventListener("resize", onWindowResize);
  window.addEventListener("keydown", onKeyDown);
  renderer.domElement.addEventListener("mousedown", onMouseDown);
  renderer.domElement.addEventListener("mousemove", onMouseMove);
  renderer.domElement.addEventListener("mouseup", onMouseUp);
  renderer.domElement.addEventListener("click", onClick);
  renderer.domElement.addEventListener("contextmenu", (e) =>
    e.preventDefault()
  ); // Prevent right-click menu
  renderer.domElement.addEventListener("wheel", onWheel, { passive: false });
  renderer.domElement.addEventListener("touchstart", onTouchStart, {
    passive: false,
  });
  renderer.domElement.addEventListener("touchmove", onTouchMove, {
    passive: false,
  });
  renderer.domElement.addEventListener("touchend", onTouchEnd);
  // Double-click disabled - using inline editing instead
  // renderer.domElement.addEventListener("dblclick", onDoubleClick);

  // Cell input handling
  const cellInput = document.getElementById("cell-input");
  cellInput.addEventListener("blur", onCellInputBlur);
  cellInput.addEventListener("keydown", onCellInputKeydown);

  // Color picker handling
  const colorPicker = document.getElementById("color-picker");
  colorPicker.addEventListener("input", onColorChange);

  // Text color picker handling
  const textColorPicker = document.getElementById("text-color-picker");
  textColorPicker.addEventListener("input", onTextColorChange);

  // Start animation loop
  animate();
}

function createCellGrid() {
  const cellGeometry = new THREE.BoxGeometry(
    CELL_WIDTH,
    CELL_HEIGHT,
    CELL_DEPTH
  );
  const edgesGeometry = new THREE.EdgesGeometry(cellGeometry);

  for (let z = 0; z < GRID_SIZE_Z; z++) {
    for (let y = 0; y < GRID_SIZE_Y; y++) {
      for (let x = 0; x < GRID_SIZE_X; x++) {
        // Create cell mesh with subtle white fill (no lighting by default)
        const cellMaterial = new THREE.MeshBasicMaterial({
          color: 0xffffff,
          transparent: true,
          opacity: 0.1, // 10% opaque for visibility
          depthWrite: false, // Don't write to depth buffer to reduce z-fighting
          side: THREE.DoubleSide,
        });
        const cellMesh = new THREE.Mesh(cellGeometry, cellMaterial);
        cellMesh.renderOrder = 1; // Render cells before edges

        // Position cell from top-left corner with offset for labels
        const posX = LABEL_OFFSET_X + x * CELL_WIDTH + CELL_WIDTH / 2;
        const posY = -LABEL_OFFSET_Y - y * CELL_HEIGHT - CELL_HEIGHT / 2;
        const posZ = z * CELL_DEPTH + CELL_DEPTH / 2;

        cellMesh.position.set(posX, posY, posZ);
        cellMesh.userData = { x, y, z, type: "cell" };

        // Create edges
        const edgeMaterial = new THREE.LineBasicMaterial({
          color: 0xd0d0d0,
          linewidth: 1,
          depthTest: true,
        });
        const edges = new THREE.LineSegments(edgesGeometry, edgeMaterial);
        edges.renderOrder = 2; // Render edges after cell fill
        cellMesh.add(edges);

        pivot.add(cellMesh);
        cellMeshes.push(cellMesh);
      }
    }
  }
}

function createLabels() {
  // Column labels (A, B, C, ...) - positioned at the TOP, aligned with cell columns
  for (let x = 0; x < GRID_SIZE_X; x++) {
    const label = columnToLetter(x);
    const posX = LABEL_OFFSET_X + x * CELL_WIDTH + CELL_WIDTH / 2; // Aligned with cell center X
    const posY = -LABEL_OFFSET_Y + CELL_HEIGHT / 2; // Just above the first row
    const posZ = 0; // At the very front edge (z=0 layer)

    const sprite = createLabelSprite(label, 64, "#666666");
    sprite.position.set(posX, posY, posZ);
    sprite.userData = { type: "label" };
    pivot.add(sprite);
    labelSprites.push(sprite);
  }

  // Row labels (1, 2, 3, ...) - positioned on the LEFT, aligned with cell rows
  for (let y = 0; y < GRID_SIZE_Y; y++) {
    const label = (y + 1).toString();
    const posX = LABEL_OFFSET_X - CELL_WIDTH / 2; // Just to the left of the first column
    const posY = -LABEL_OFFSET_Y - y * CELL_HEIGHT - CELL_HEIGHT / 2; // Aligned with cell center Y
    const posZ = 0; // At the very front edge (z=0 layer)

    const sprite = createLabelSprite(label, 64, "#666666");
    sprite.position.set(posX, posY, posZ);
    sprite.userData = { type: "label" };
    pivot.add(sprite);
    labelSprites.push(sprite);
  }

  // Z-axis labels (I, II, III, IV, V...) - positioned at the BOTTOM, aligned with depth layers
  for (let z = 0; z < GRID_SIZE_Z; z++) {
    const label = toRomanNumeral(z);
    const posX = LABEL_OFFSET_X - CELL_WIDTH / 2; // Aligned with row labels on the left
    const posY = -LABEL_OFFSET_Y - GRID_SIZE_Y * CELL_HEIGHT - CELL_HEIGHT / 2; // Below the last row
    const posZ = z * CELL_DEPTH + CELL_DEPTH / 2; // Aligned with each Z layer

    const sprite = createLabelSprite(label, 64, "#666666");
    sprite.position.set(posX, posY, posZ);
    sprite.userData = { type: "label" };
    pivot.add(sprite);
    labelSprites.push(sprite);
  }
}

function columnToLetter(index) {
  let letter = "";
  let num = index;
  while (num >= 0) {
    letter = String.fromCharCode(65 + (num % 26)) + letter;
    num = Math.floor(num / 26) - 1;
  }
  return letter;
}

function toRomanNumeral(num) {
  const romanNumerals = [
    { value: 10, symbol: "X" },
    { value: 9, symbol: "IX" },
    { value: 5, symbol: "V" },
    { value: 4, symbol: "IV" },
    { value: 1, symbol: "I" },
  ];

  let result = "";
  let n = num + 1; // Convert 0-indexed to 1-indexed

  for (let i = 0; i < romanNumerals.length; i++) {
    while (n >= romanNumerals[i].value) {
      result += romanNumerals[i].symbol;
      n -= romanNumerals[i].value;
    }
  }

  return result;
}

function createLabelSprite(text, fontSize = 64, color = "#000000") {
  const canvas = document.createElement("canvas");
  const context = canvas.getContext("2d");

  // Set canvas size
  canvas.width = 256;
  canvas.height = 128;

  // Configure text
  context.font = `${fontSize}px -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif`;
  context.fillStyle = color;
  context.textAlign = "center";
  context.textBaseline = "middle";

  // Draw text
  context.fillText(text, canvas.width / 2, canvas.height / 2);

  // Create texture
  const texture = new THREE.CanvasTexture(canvas);
  texture.needsUpdate = true;

  // Create sprite
  const spriteMaterial = new THREE.SpriteMaterial({ map: texture });
  const sprite = new THREE.Sprite(spriteMaterial);
  sprite.scale.set(70, 35, 1); // Fixed size for labels, larger and more visible

  return sprite;
}

// Helper function to get font family from selector value
function getFontFamily(fontName) {
  const fontMap = {
    system: "-apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif",
    arial: "Arial, sans-serif",
    helvetica: "Helvetica, Arial, sans-serif",
    times: "'Times New Roman', Times, serif",
    georgia: "Georgia, serif",
    courier: "'Courier New', Courier, monospace",
    monaco: "Monaco, 'Courier New', monospace",
    verdana: "Verdana, Geneva, sans-serif",
    comic: "'Comic Sans MS', cursive",
    impact: "Impact, Charcoal, sans-serif",
  };
  return fontMap[fontName] || fontMap.arial;
}

function createTextSprite(
  text,
  fontSize = 64,
  color = "#000000",
  bold = false,
  italic = false,
  strikethrough = false,
  fontFamily = "arial"
) {
  const canvas = document.createElement("canvas");
  const context = canvas.getContext("2d");

  // Set canvas size
  canvas.width = 512;
  canvas.height = 256;

  // Get font family string
  const fontFamilyStr = getFontFamily(fontFamily);

  // Configure text with formatting
  let fontStyle = "";
  if (italic) fontStyle += "italic ";
  if (bold) fontStyle += "bold ";

  context.font = `${fontStyle}${fontSize}px ${fontFamilyStr}`;
  context.fillStyle = color;
  context.textAlign = "center";
  context.textBaseline = "middle";

  // Draw text
  context.fillText(text, canvas.width / 2, canvas.height / 2);

  // Draw strikethrough line if enabled
  if (strikethrough) {
    const textMetrics = context.measureText(text);
    const textWidth = textMetrics.width;
    const centerX = canvas.width / 2;
    const centerY = canvas.height / 2;

    context.strokeStyle = color;
    context.lineWidth = fontSize / 20;
    context.beginPath();
    context.moveTo(centerX - textWidth / 2, centerY);
    context.lineTo(centerX + textWidth / 2, centerY);
    context.stroke();
  }

  // Create texture
  const texture = new THREE.CanvasTexture(canvas);
  texture.needsUpdate = true;

  // Create sprite
  const spriteMaterial = new THREE.SpriteMaterial({
    map: texture,
    depthTest: false, // Render on top
    depthWrite: false,
  });
  const sprite = new THREE.Sprite(spriteMaterial);
  sprite.scale.set(CELL_WIDTH * 0.8, CELL_HEIGHT * 1.2, 1); // Readable size
  sprite.renderOrder = 999; // Render after everything else

  return sprite;
}

function updateCellText(x, y, z, text) {
  const key = `${x},${y},${z}`;

  // Remove old sprite if exists
  const oldSprite = textSprites.find(
    (s) => s.userData.x === x && s.userData.y === y && s.userData.z === z
  );
  if (oldSprite) {
    pivot.remove(oldSprite);
    textSprites = textSprites.filter((s) => s !== oldSprite);
  }

  // Find the cell mesh
  const cellMesh = cellMeshes.find(
    (mesh) =>
      mesh.userData.x === x && mesh.userData.y === y && mesh.userData.z === z
  );

  if (text && text.trim() !== "") {
    // Store data
    cellData[key] = text;

    // Get text color (default to black if not set)
    const textColor = cellTextColors[key] || "#000000";

    // Get text formatting states
    const isBold = cellTextBold[key] || false;
    const isItalic = cellTextItalic[key] || false;
    const isStrikethrough = cellTextStrikethrough[key] || false;
    const fontFamily = cellFontFamily[key] || "arial";

    // Create new sprite from top-left with offset for labels
    const sprite = createTextSprite(
      text,
      100,
      textColor,
      isBold,
      isItalic,
      isStrikethrough,
      fontFamily
    );
    const posX = LABEL_OFFSET_X + x * CELL_WIDTH + CELL_WIDTH / 2;
    const posY = -LABEL_OFFSET_Y - y * CELL_HEIGHT - CELL_HEIGHT / 2;
    const posZ = z * CELL_DEPTH + CELL_DEPTH / 2;

    sprite.position.set(posX, posY, posZ);
    sprite.userData = { x, y, z, type: "text" };

    pivot.add(sprite);
    textSprites.push(sprite);

    // Update cell visual: white background + thicker border
    if (cellMesh) {
      // Check if cell has custom background color
      const customColor = cellBackgroundColors[key];
      if (customColor) {
        // Use lighting material for colored cells
        cellMesh.material.dispose();
        cellMesh.material = new THREE.MeshStandardMaterial({
          color: new THREE.Color(customColor),
          transparent: true,
          opacity: 1.0,
          depthWrite: false,
          side: THREE.DoubleSide,
          metalness: 0.0,
          roughness: 0.5,
          emissive: new THREE.Color(customColor),
          emissiveIntensity: 0.15,
        });
        cellMesh.renderOrder = 5; // Colored cells render on top
      } else {
        // Use basic material for non-colored cells
        cellMesh.material.dispose();
        cellMesh.material = new THREE.MeshBasicMaterial({
          color: 0xffffff,
          transparent: true,
          opacity: 0.2,
          depthWrite: false,
          side: THREE.DoubleSide,
        });
        cellMesh.renderOrder = 3; // Cells with text render above empty cells
      }

      // Update the edge color to be darker
      const edges = cellMesh.children.find(
        (child) => child.type === "LineSegments"
      );
      if (edges) {
        edges.material.color.setHex(0x333333);
      }

      // Add thick border if not already present
      if (!cellMesh.userData.thickBorder) {
        const cellGeometry = new THREE.BoxGeometry(
          CELL_WIDTH,
          CELL_HEIGHT,
          CELL_DEPTH
        );
        const edgesGeometry = new THREE.EdgesGeometry(cellGeometry);

        // Create a simpler thick border with just 3 overlapping lines
        const thickBorderGroup = new THREE.Group();
        const offsets = [0, 0.3, -0.3];
        offsets.forEach((offset) => {
          const borderMaterial = new THREE.LineBasicMaterial({
            color: 0x333333,
            transparent: false, // Disable transparency to avoid z-fighting
            depthWrite: true,
          });
          const borderEdges = new THREE.LineSegments(
            edgesGeometry,
            borderMaterial
          );
          borderEdges.position.set(offset, offset, offset);
          thickBorderGroup.add(borderEdges);
        });

        thickBorderGroup.userData = { type: "thickBorder" };
        cellMesh.add(thickBorderGroup);
        cellMesh.userData.thickBorder = thickBorderGroup;
      }
    }
  } else {
    // Remove data
    delete cellData[key];

    // Restore cell visual: subtle transparent + lighter border
    if (cellMesh) {
      // Check if cell has custom background color
      const customColor = cellBackgroundColors[key];
      if (customColor) {
        // Use lighting material for colored cells
        cellMesh.material.dispose();
        cellMesh.material = new THREE.MeshStandardMaterial({
          color: new THREE.Color(customColor),
          transparent: true,
          opacity: 1.0,
          depthWrite: false,
          side: THREE.DoubleSide,
          metalness: 0.0,
          roughness: 0.5,
          emissive: new THREE.Color(customColor),
          emissiveIntensity: 0.15,
        });
        cellMesh.renderOrder = 5; // Keep colored cells on top
      } else {
        // Use basic material for non-colored cells
        cellMesh.material.dispose();
        cellMesh.material = new THREE.MeshBasicMaterial({
          color: 0xffffff,
          transparent: true,
          opacity: 0.1,
          depthWrite: false,
          side: THREE.DoubleSide,
        });
        cellMesh.renderOrder = 1; // Back to default render order
      }

      // Restore the edge color to default
      const edges = cellMesh.children.find(
        (child) => child.type === "LineSegments"
      );
      if (edges) {
        edges.material.color.setHex(0xd0d0d0);
      }

      // Remove thick border if present
      if (cellMesh.userData.thickBorder) {
        cellMesh.remove(cellMesh.userData.thickBorder);
        cellMesh.userData.thickBorder = null;
      }
    }
  }
}

function onKeyDown(event) {
  // Ignore if meta/ctrl keys are pressed (for shortcuts like Cmd+R)
  if (event.metaKey || event.ctrlKey) {
    return;
  }

  // Handle arrow keys for navigation (works both in and out of edit mode)
  if (
    event.key === "ArrowLeft" ||
    event.key === "ArrowRight" ||
    event.key === "ArrowUp" ||
    event.key === "ArrowDown"
  ) {
    // If editing, finish first
    if (isEditingCell) {
      finishEditing(true);
    }

    if (selectionStart) {
      // Determine direction
      let deltaX = 0;
      let deltaY = 0;
      let deltaZ = 0;

      if (event.altKey) {
        // Option/Alt key + Up/Down moves in Z direction
        if (event.key === "ArrowUp") {
          deltaZ = -1; // Move towards front (lower Z)
        } else if (event.key === "ArrowDown") {
          deltaZ = 1; // Move towards back (higher Z)
        }
      } else {
        // Regular arrow keys move in X/Y plane
        if (event.key === "ArrowLeft") {
          deltaX = -1;
        } else if (event.key === "ArrowRight") {
          deltaX = 1;
        } else if (event.key === "ArrowUp") {
          deltaY = -1;
        } else if (event.key === "ArrowDown") {
          deltaY = 1;
        }
      }

      // Get the reference cell (use selectionEnd for extending, selectionStart for moving)
      const referenceCell = event.shiftKey ? selectionEnd : selectionStart;
      const newX = Math.max(
        0,
        Math.min(GRID_SIZE_X - 1, referenceCell.x + deltaX)
      );
      const newY = Math.max(
        0,
        Math.min(GRID_SIZE_Y - 1, referenceCell.y + deltaY)
      );
      const newZ = Math.max(
        0,
        Math.min(GRID_SIZE_Z - 1, referenceCell.z + deltaZ)
      );

      // Find the target cell
      const targetCell = cellMeshes.find(
        (mesh) =>
          mesh.userData.x === newX &&
          mesh.userData.y === newY &&
          mesh.userData.z === newZ
      );

      if (targetCell) {
        if (event.shiftKey) {
          // Extend selection
          selectionEnd = targetCell.userData;
        } else {
          // Move selection
          selectionStart = targetCell.userData;
          selectionEnd = targetCell.userData;
        }
        selectCubicRegion(selectionStart, selectionEnd);
      }
    }

    event.preventDefault();
    return;
  }

  // If we're editing, handle edit keys
  if (isEditingCell && editingCellCoords) {
    if (event.key === "Enter") {
      // Store current coordinates before finishing edit
      const currentX = editingCellCoords.x;
      const currentY = editingCellCoords.y;
      const currentZ = editingCellCoords.z;

      console.log(
        "Enter pressed, current cell:",
        currentX,
        currentY,
        currentZ,
        "GRID_SIZE_Y:",
        GRID_SIZE_Y
      );

      // Save and exit edit mode
      finishEditing(true);

      // Check if there's a cell below
      if (currentY + 1 < GRID_SIZE_Y) {
        console.log(
          "Looking for cell below at:",
          currentX,
          currentY + 1,
          currentZ
        );

        // Find the cell below in cellMeshes
        const cellBelow = cellMeshes.find(
          (mesh) =>
            mesh.userData.x === currentX &&
            mesh.userData.y === currentY + 1 &&
            mesh.userData.z === currentZ
        );

        console.log("Cell below found:", cellBelow ? "YES" : "NO");

        if (cellBelow) {
          // Select the cell below
          selectionStart = cellBelow.userData;
          selectionEnd = cellBelow.userData;
          selectedCellMesh = cellBelow;
          selectCubicRegion(selectionStart, selectionEnd);

          // Start editing the cell below
          startEditing(currentX, currentY + 1, currentZ);
        }
      } else {
        console.log("Already at bottom row, currentY + 1:", currentY + 1);
      }

      event.preventDefault();
      return;
    } else if (event.key === "Escape") {
      // Cancel edit
      finishEditing(false);
      event.preventDefault();
      return;
    } else if (event.key === "Backspace") {
      // Clear all selected cells and exit edit mode
      finishEditing(false); // Don't save current edits

      if (selectionStart && selectionEnd) {
        const minX = Math.min(selectionStart.x, selectionEnd.x);
        const maxX = Math.max(selectionStart.x, selectionEnd.x);
        const minY = Math.min(selectionStart.y, selectionEnd.y);
        const maxY = Math.max(selectionStart.y, selectionEnd.y);
        const minZ = Math.min(selectionStart.z, selectionEnd.z);
        const maxZ = Math.max(selectionStart.z, selectionEnd.z);

        // Clear all cells in the range
        for (let x = minX; x <= maxX; x++) {
          for (let y = minY; y <= maxY; y++) {
            for (let z = minZ; z <= maxZ; z++) {
              updateCellText(x, y, z, "");
            }
          }
        }
      }
      event.preventDefault();
      return;
    } else if (event.key.length === 1) {
      // Add character
      editingText += event.key;
      updateCellText(
        editingCellCoords.x,
        editingCellCoords.y,
        editingCellCoords.z,
        editingText
      );
      event.preventDefault();
      return;
    }
    return;
  }

  // If a cell is selected but not editing yet
  if (selectionStart && !isEditingCell) {
    // Press Delete to clear selected cells
    if (event.key === "Delete" || event.key === "Backspace") {
      if (selectionStart && selectionEnd) {
        const minX = Math.min(selectionStart.x, selectionEnd.x);
        const maxX = Math.max(selectionStart.x, selectionEnd.x);
        const minY = Math.min(selectionStart.y, selectionEnd.y);
        const maxY = Math.max(selectionStart.y, selectionEnd.y);
        const minZ = Math.min(selectionStart.z, selectionEnd.z);
        const maxZ = Math.max(selectionStart.z, selectionEnd.z);

        // Clear all cells in the range
        for (let x = minX; x <= maxX; x++) {
          for (let y = minY; y <= maxY; y++) {
            for (let z = minZ; z <= maxZ; z++) {
              updateCellText(x, y, z, "");
            }
          }
        }
      }
      event.preventDefault();
      return;
    }

    // Any other key starts editing
    if (event.key.length === 1 || event.key === "Enter") {
      startEditing(selectionStart.x, selectionStart.y, selectionStart.z);
      if (event.key.length === 1) {
        // Start with the typed character
        editingText = event.key;
        updateCellText(
          editingCellCoords.x,
          editingCellCoords.y,
          editingCellCoords.z,
          editingText
        );
      }
      event.preventDefault();
    }
  }
}

function startEditing(x, y, z) {
  console.log(
    `startEditing called for cell (${x},${y},${z}), isEditingCell was:`,
    isEditingCell
  );
  isEditingCell = true;
  editingCellCoords = { x, y, z };
  const key = `${x},${y},${z}`;
  editingText = cellData[key] || "";
  console.log("Edit mode started, editingText:", editingText);

  // Show current text (or empty if no text)
  // Don't clear it - keep existing value visible
}

function finishEditing(save) {
  console.log(
    "finishEditing called, save:",
    save,
    "isEditingCell:",
    isEditingCell,
    "editingCellCoords:",
    editingCellCoords
  );
  if (!isEditingCell || !editingCellCoords) {
    console.log("finishEditing early return");
    return;
  }

  if (save) {
    // Text is already updated in real-time, just clean up
    if (editingText.trim() === "") {
      updateCellText(
        editingCellCoords.x,
        editingCellCoords.y,
        editingCellCoords.z,
        ""
      );
    }
  } else {
    // Restore original value
    const key = `${editingCellCoords.x},${editingCellCoords.y},${editingCellCoords.z}`;
    const originalValue = cellData[key] || "";
    updateCellText(
      editingCellCoords.x,
      editingCellCoords.y,
      editingCellCoords.z,
      originalValue
    );
  }

  isEditingCell = false;
  editingCellCoords = null;
  editingText = "";
  console.log("finishEditing complete, isEditingCell now:", isEditingCell);
}

function getMouseCoordinates(event) {
  // Adjust for toolbar offset - canvas starts TOOLBAR_HEIGHT pixels from top
  const canvasHeight = window.innerHeight - TOOLBAR_HEIGHT;
  const adjustedY = event.clientY - TOOLBAR_HEIGHT;

  const mouse = new THREE.Vector2();
  mouse.x = (event.clientX / window.innerWidth) * 2 - 1;
  mouse.y = -(adjustedY / canvasHeight) * 2 + 1;

  return mouse;
}

function onWindowResize() {
  // Update grid size based on new window dimensions
  GRID_SIZE_X = Math.floor((window.innerWidth - LABEL_OFFSET_X) / CELL_WIDTH);
  GRID_SIZE_Y = Math.floor(
    (window.innerHeight - TOOLBAR_HEIGHT - LABEL_OFFSET_Y) / CELL_HEIGHT
  );

  const gridWidth = GRID_SIZE_X * CELL_WIDTH;
  const gridHeight = GRID_SIZE_Y * CELL_HEIGHT;
  const centerX = gridWidth / 2;
  const centerY = -gridHeight / 2;

  // Update orthographic camera for new grid size
  orthographicCamera.left = -gridWidth / 2;
  orthographicCamera.right = gridWidth / 2;
  orthographicCamera.top = gridHeight / 2;
  orthographicCamera.bottom = -gridHeight / 2;
  orthographicCamera.updateProjectionMatrix();
  orthographicCamera.position.set(centerX, centerY, 3000);
  orthographicCamera.lookAt(centerX, centerY, 0);

  // Update perspective camera for new aspect ratio
  perspectiveCamera.aspect =
    window.innerWidth / (window.innerHeight - TOOLBAR_HEIGHT);
  perspectiveCamera.updateProjectionMatrix();
  perspectiveCamera.position.set(centerX, centerY, 3000);
  perspectiveCamera.lookAt(centerX, centerY, 0);

  renderer.setSize(window.innerWidth, window.innerHeight - TOOLBAR_HEIGHT);
  anaglyphEffect.setSize(
    window.innerWidth,
    window.innerHeight - TOOLBAR_HEIGHT
  );
}

function onMouseDown(event) {
  // Check if right mouse button is pressed for panning
  if (event.button === 2) {
    isPanning = true;
    previousPanPosition = {
      x: event.clientX,
      y: event.clientY,
    };
    renderer.domElement.classList.add("grabbing");
    event.preventDefault();
    return;
  }

  // Check if Cmd key (Meta key on Mac) is pressed for rotation
  if (event.metaKey) {
    isRotating = true;
    previousMousePosition = {
      x: event.clientX,
      y: event.clientY,
    };
    renderer.domElement.classList.add("grabbing");
    event.preventDefault();
    return;
  }

  // Start drag selection (if not shift+clicking)
  if (!event.shiftKey) {
    const raycaster = new THREE.Raycaster();
    const mouse = getMouseCoordinates(event);

    raycaster.setFromCamera(mouse, camera);
    const intersects = raycaster.intersectObjects(cellMeshes, true);

    if (intersects.length > 0) {
      const intersect = intersects[0];
      let cellMesh = intersect.object;

      while (cellMesh && !cellMesh.userData.type) {
        cellMesh = cellMesh.parent;
      }

      if (cellMesh && cellMesh.userData.type === "cell") {
        isDragging = true;
        actuallyDragged = false; // Reset the flag
        selectionStart = cellMesh.userData;
        selectionEnd = cellMesh.userData;
        selectCubicRegion(selectionStart, selectionEnd);
      }
    }
  }
}

function onMouseMove(event) {
  if (isPanning) {
    const deltaX = event.clientX - previousPanPosition.x;
    const deltaY = event.clientY - previousPanPosition.y;

    // Pan the camera by adjusting its position
    // Scale the pan speed based on the zoom level
    const panSpeed = 2.0 / camera.zoom;

    camera.position.x -= deltaX * panSpeed;
    camera.position.y += deltaY * panSpeed; // Inverted because screen Y goes down

    previousPanPosition = {
      x: event.clientX,
      y: event.clientY,
    };

    event.preventDefault();
  } else if (isRotating && event.metaKey) {
    const deltaX = event.clientX - previousMousePosition.x;
    const deltaY = event.clientY - previousMousePosition.y;

    // Rotate around Y axis (horizontal mouse movement)
    pivot.rotation.y += deltaX * 0.01;

    // Rotate around X axis (vertical mouse movement)
    pivot.rotation.x += deltaY * 0.01;

    previousMousePosition = {
      x: event.clientX,
      y: event.clientY,
    };

    event.preventDefault();
  } else if (isDragging) {
    // Update selection end point during drag
    const raycaster = new THREE.Raycaster();
    const mouse = getMouseCoordinates(event);

    raycaster.setFromCamera(mouse, camera);
    const intersects = raycaster.intersectObjects(cellMeshes, true);

    if (intersects.length > 0) {
      const intersect = intersects[0];
      let cellMesh = intersect.object;

      while (cellMesh && !cellMesh.userData.type) {
        cellMesh = cellMesh.parent;
      }

      if (cellMesh && cellMesh.userData.type === "cell") {
        // Check if we moved to a different cell
        if (
          selectionEnd.x !== cellMesh.userData.x ||
          selectionEnd.y !== cellMesh.userData.y ||
          selectionEnd.z !== cellMesh.userData.z
        ) {
          actuallyDragged = true;
        }
        selectionEnd = cellMesh.userData;
        selectCubicRegion(selectionStart, selectionEnd);
      }
    }
  }
}

function onMouseUp(event) {
  if (isPanning) {
    isPanning = false;
    renderer.domElement.classList.remove("grabbing");
  }

  if (isRotating) {
    isRotating = false;
    renderer.domElement.classList.remove("grabbing");
  }

  // Finish drag selection
  if (isDragging) {
    isDragging = false;
    // Only set justFinishedDragging if we actually dragged to a different cell
    if (actuallyDragged) {
      justFinishedDragging = true;
      // Reset flag after click event has been processed
      setTimeout(() => {
        justFinishedDragging = false;
      }, 50);
    }
    actuallyDragged = false; // Reset for next time
  }
}

function onClick(event) {
  console.log(
    "onClick fired! isEditingCell:",
    isEditingCell,
    "isDragging:",
    isDragging,
    "justFinishedDragging:",
    justFinishedDragging
  );

  // Don't select if we were rotating or just finished dragging
  if (event.metaKey || isDragging || justFinishedDragging) {
    console.log("onClick returning early");
    return;
  }

  // Finish any current editing before selecting new cell
  if (isEditingCell) {
    console.log("Finishing edit before selecting new cell");
    finishEditing(true);
  }

  const raycaster = new THREE.Raycaster();
  const mouse = getMouseCoordinates(event);

  raycaster.setFromCamera(mouse, camera);

  // Check intersection with cells
  const intersects = raycaster.intersectObjects(cellMeshes, true);

  if (intersects.length > 0) {
    const intersect = intersects[0];
    let cellMesh = intersect.object;

    // Find the parent cell mesh
    while (cellMesh && !cellMesh.userData.type) {
      cellMesh = cellMesh.parent;
    }

    if (cellMesh && cellMesh.userData.type === "cell") {
      if (event.shiftKey && selectionStart) {
        // Shift+click to extend selection - don't enter edit mode
        selectionEnd = cellMesh.userData;
        selectCubicRegion(selectionStart, selectionEnd);
      } else {
        // Regular click - single cell selection (don't auto-edit to allow arrow key navigation)
        selectionStart = cellMesh.userData;
        selectionEnd = cellMesh.userData;
        selectCubicRegion(selectionStart, selectionEnd);
      }
    }
  } else {
    // Clicked on empty space, deselect and stop editing
    if (isEditingCell) {
      finishEditing(true);
    }
    deselectCell();
  }
}

function selectCubicRegion(start, end) {
  if (!start || !end) return;

  // Remove previous selection outline only (keep selection range)
  if (selectionOutline) {
    pivot.remove(selectionOutline);
    selectionOutline = null;
  }

  // Calculate the bounding box of the selection
  const minX = Math.min(start.x, end.x);
  const maxX = Math.max(start.x, end.x);
  const minY = Math.min(start.y, end.y);
  const maxY = Math.max(start.y, end.y);
  const minZ = Math.min(start.z, end.z);
  const maxZ = Math.max(start.z, end.z);

  const width = (maxX - minX + 1) * CELL_WIDTH;
  const height = (maxY - minY + 1) * CELL_HEIGHT;
  const depth = (maxZ - minZ + 1) * CELL_DEPTH;

  // Calculate center position of the selection box (center between first and last cell centers)
  const centerX =
    LABEL_OFFSET_X + ((minX + maxX) * CELL_WIDTH) / 2 + CELL_WIDTH / 2;
  const centerY =
    -LABEL_OFFSET_Y - ((minY + maxY) * CELL_HEIGHT) / 2 - CELL_HEIGHT / 2;
  const centerZ = ((minZ + maxZ) * CELL_DEPTH) / 2 + CELL_DEPTH / 2;

  // Create a group for the selection
  selectionOutline = new THREE.Group();

  // Create the cubic selection box
  const selectionGeometry = new THREE.BoxGeometry(width, height, depth);

  // Create simpler edge outlines to avoid z-fighting
  const offsets = [0, 0.3, -0.3];
  offsets.forEach((offset) => {
    const edgesGeometry = new THREE.EdgesGeometry(selectionGeometry);
    const selectionMaterial = new THREE.LineBasicMaterial({
      color: 0x1a74e8,
      transparent: false,
      depthTest: false, // Always render on top, ignore depth
    });

    const edges = new THREE.LineSegments(edgesGeometry, selectionMaterial);
    edges.position.set(offset, offset, offset);
    edges.renderOrder = 100; // Render selection edges above everything
    selectionOutline.add(edges);
  });

  // Add semi-transparent fill for better visibility
  const fillMaterial = new THREE.MeshBasicMaterial({
    color: 0x1a74e8,
    transparent: true,
    opacity: 0.1,
    depthWrite: false,
    depthTest: false, // Always render on top
    side: THREE.DoubleSide,
  });
  const fillMesh = new THREE.Mesh(selectionGeometry, fillMaterial);
  fillMesh.renderOrder = 50; // Render fill above cells but below edges
  selectionOutline.add(fillMesh);

  selectionOutline.position.set(centerX, centerY, centerZ);
  pivot.add(selectionOutline);
}

function selectCell(cellMesh) {
  // Remove previous selection
  deselectCell();

  // Store selected cell
  selectedCellMesh = cellMesh;
  const { x, y, z } = cellMesh.userData;

  // Create a group for the selection
  selectionOutline = new THREE.Group();

  // Create thicker outline using multiple edge lines
  const cellGeometry = new THREE.BoxGeometry(
    CELL_WIDTH,
    CELL_HEIGHT,
    CELL_DEPTH
  );

  // Create simpler edge outlines to avoid z-fighting
  const offsets = [0, 0.3, -0.3];
  offsets.forEach((offset) => {
    const edgesGeometry = new THREE.EdgesGeometry(cellGeometry);
    const selectionMaterial = new THREE.LineBasicMaterial({
      color: 0x1a74e8,
      transparent: false,
      depthTest: false, // Always render on top, ignore depth
    });

    const edges = new THREE.LineSegments(edgesGeometry, selectionMaterial);
    edges.position.set(offset, offset, offset);
    edges.renderOrder = 100; // Render selection edges above everything
    selectionOutline.add(edges);
  });

  // Add semi-transparent fill for better visibility
  const fillMaterial = new THREE.MeshBasicMaterial({
    color: 0x1a74e8,
    transparent: true,
    opacity: 0.1,
    depthWrite: false,
    depthTest: false, // Always render on top
    side: THREE.DoubleSide,
  });
  const fillMesh = new THREE.Mesh(cellGeometry, fillMaterial);
  fillMesh.renderOrder = 50; // Render fill above cells but below edges
  selectionOutline.add(fillMesh);

  selectionOutline.position.copy(cellMesh.position);
  pivot.add(selectionOutline);
}

function deselectCell() {
  if (selectionOutline) {
    pivot.remove(selectionOutline);
    selectionOutline = null;
  }
  selectionStart = null;
  selectionEnd = null;
  selectedCellMesh = null;
}

function onDoubleClick(event) {
  const raycaster = new THREE.Raycaster();
  const mouse = getMouseCoordinates(event);

  raycaster.setFromCamera(mouse, camera);

  // Check intersection with cells
  const intersects = raycaster.intersectObjects(cellMeshes, true);

  if (intersects.length > 0) {
    const intersect = intersects[0];
    let cellMesh = intersect.object;

    // Find the parent cell mesh
    while (cellMesh && !cellMesh.userData.type) {
      cellMesh = cellMesh.parent;
    }

    if (cellMesh && cellMesh.userData.type === "cell") {
      const { x, y, z } = cellMesh.userData;
      openCellEditor(x, y, z, intersect.point);
    }
  }
}

function openCellEditor(x, y, z, worldPosition) {
  selectedCell = { x, y, z };

  // Convert 3D position to screen position
  const vector = worldPosition.clone();
  vector.project(camera);

  const screenX = ((vector.x + 1) / 2) * window.innerWidth;
  const screenY = (-(vector.y - 1) / 2) * window.innerHeight;

  // Get current value
  const key = `${x},${y},${z}`;
  const currentValue = cellData[key] || "";

  // Position and show input
  const cellInput = document.getElementById("cell-input");
  cellInput.style.left = screenX + "px";
  cellInput.style.top = screenY + "px";
  cellInput.style.display = "block";
  cellInput.value = currentValue;
  cellInput.focus();
  cellInput.select();
}

function onCellInputBlur() {
  if (selectedCell) {
    const cellInput = document.getElementById("cell-input");
    const value = cellInput.value;

    updateCellText(selectedCell.x, selectedCell.y, selectedCell.z, value);

    cellInput.style.display = "none";
    selectedCell = null;
  }
}

function onCellInputKeydown(event) {
  if (event.key === "Enter") {
    event.target.blur();
  } else if (event.key === "Escape") {
    const cellInput = document.getElementById("cell-input");
    cellInput.style.display = "none";
    selectedCell = null;
  }
}

function onColorChange(event) {
  const color = event.target.value;

  // Apply color to all selected cells
  if (selectionStart && selectionEnd) {
    const minX = Math.min(selectionStart.x, selectionEnd.x);
    const maxX = Math.max(selectionStart.x, selectionEnd.x);
    const minY = Math.min(selectionStart.y, selectionEnd.y);
    const maxY = Math.max(selectionStart.y, selectionEnd.y);
    const minZ = Math.min(selectionStart.z, selectionEnd.z);
    const maxZ = Math.max(selectionStart.z, selectionEnd.z);

    // Set background color for all cells in the range
    for (let x = minX; x <= maxX; x++) {
      for (let y = minY; y <= maxY; y++) {
        for (let z = minZ; z <= maxZ; z++) {
          setCellBackgroundColor(x, y, z, color);
        }
      }
    }
  }
}

function onTextColorChange(event) {
  const color = event.target.value;

  // Apply text color to all selected cells
  if (selectionStart && selectionEnd) {
    const minX = Math.min(selectionStart.x, selectionEnd.x);
    const maxX = Math.max(selectionStart.x, selectionEnd.x);
    const minY = Math.min(selectionStart.y, selectionEnd.y);
    const maxY = Math.max(selectionStart.y, selectionEnd.y);
    const minZ = Math.min(selectionStart.z, selectionEnd.z);
    const maxZ = Math.max(selectionStart.z, selectionEnd.z);

    // Set text color for all cells in the range
    for (let x = minX; x <= maxX; x++) {
      for (let y = minY; y <= maxY; y++) {
        for (let z = minZ; z <= maxZ; z++) {
          setCellTextColor(x, y, z, color);
        }
      }
    }
  }
}

function setCellBackgroundColor(x, y, z, color) {
  const key = `${x},${y},${z}`;
  cellBackgroundColors[key] = color;

  // Find the cell mesh
  const cellMesh = cellMeshes.find(
    (mesh) =>
      mesh.userData.x === x && mesh.userData.y === y && mesh.userData.z === z
  );

  if (cellMesh) {
    // Convert hex color to THREE color
    const threeColor = new THREE.Color(color);

    // Replace material with lighting-enabled material for colored cells
    cellMesh.material.dispose(); // Clean up old material
    cellMesh.material = new THREE.MeshStandardMaterial({
      color: threeColor,
      transparent: true,
      opacity: 1.0, // Completely opaque
      depthWrite: false,
      side: THREE.DoubleSide,
      metalness: 0.0, // No metallic for pure colors
      roughness: 0.5, // Less matte for more vibrant colors
      emissive: threeColor,
      emissiveIntensity: 0.15, // Slight self-illumination for vibrancy
    });

    // Render cells with background colors on top of default cells
    cellMesh.renderOrder = 5; // Higher than default cells (1) but lower than selection (10)
  }
}

function setCellTextColor(x, y, z, color) {
  const key = `${x},${y},${z}`;
  cellTextColors[key] = color;

  // Update the text sprite if it exists
  const textSprite = textSprites.find(
    (s) => s.userData.x === x && s.userData.y === y && s.userData.z === z
  );

  if (textSprite) {
    // Recreate the sprite with the new color
    const text = cellData[key];
    if (text) {
      updateCellText(x, y, z, text);
    }
  }
}

function autoSum() {
  // Check if we have a selection
  if (!selectionStart || !selectionEnd) {
    return;
  }

  const minX = Math.min(selectionStart.x, selectionEnd.x);
  const maxX = Math.max(selectionStart.x, selectionEnd.x);
  const minY = Math.min(selectionStart.y, selectionEnd.y);
  const maxY = Math.max(selectionStart.y, selectionEnd.y);
  const minZ = Math.min(selectionStart.z, selectionEnd.z);
  const maxZ = Math.max(selectionStart.z, selectionEnd.z);

  // Sum all numeric values in the selection
  let sum = 0;
  for (let x = minX; x <= maxX; x++) {
    for (let y = minY; y <= maxY; y++) {
      for (let z = minZ; z <= maxZ; z++) {
        const key = `${x},${y},${z}`;
        const value = cellData[key];
        if (value) {
          const numValue = parseFloat(value);
          if (!isNaN(numValue)) {
            sum += numValue;
          }
        }
      }
    }
  }

  // Find the cell directly below the selection (use selectionStart's X and Z)
  const targetY = maxY + 1;
  const targetX = selectionStart.x;
  const targetZ = selectionStart.z;

  // Check if target cell exists in grid
  if (targetY < GRID_SIZE_Y) {
    // Update the cell with the sum
    updateCellText(targetX, targetY, targetZ, sum.toString());

    // Select the target cell
    const targetCell = cellMeshes.find(
      (mesh) =>
        mesh.userData.x === targetX &&
        mesh.userData.y === targetY &&
        mesh.userData.z === targetZ
    );

    if (targetCell) {
      selectionStart = targetCell.userData;
      selectionEnd = targetCell.userData;
      selectCubicRegion(selectionStart, selectionEnd);
    }
  }
}

// Expose autoSum to window for HTML to access
window.autoSum = autoSum;

function onWheel(event) {
  event.preventDefault();

  // Zoom with mouse wheel or trackpad pinch
  const delta = event.deltaY;
  const zoomSpeed = 0.003; // Increased for more responsiveness

  // Adjust camera zoom
  camera.zoom -= delta * zoomSpeed;
  camera.zoom = Math.max(0.1, Math.min(5, camera.zoom)); // Clamp between 0.1x and 5x
  camera.updateProjectionMatrix();
}

function onTouchStart(event) {
  if (event.touches.length === 2) {
    event.preventDefault();

    // Calculate initial distance between two fingers
    const touch1 = event.touches[0];
    const touch2 = event.touches[1];
    const dx = touch2.clientX - touch1.clientX;
    const dy = touch2.clientY - touch1.clientY;
    initialPinchDistance = Math.sqrt(dx * dx + dy * dy);

    lastTouches = Array.from(event.touches);
  }
}

function onTouchMove(event) {
  if (event.touches.length === 2) {
    event.preventDefault();

    // Calculate current distance between two fingers
    const touch1 = event.touches[0];
    const touch2 = event.touches[1];
    const dx = touch2.clientX - touch1.clientX;
    const dy = touch2.clientY - touch1.clientY;
    const currentDistance = Math.sqrt(dx * dx + dy * dy);

    if (initialPinchDistance > 0) {
      // Calculate zoom delta with increased sensitivity
      const distanceChange = currentDistance - initialPinchDistance;
      const zoomDelta = distanceChange * 0.005; // Increased sensitivity

      // Apply zoom incrementally for smoother response
      camera.zoom += zoomDelta;
      camera.zoom = Math.max(0.1, Math.min(5, camera.zoom)); // Clamp between 0.1x and 5x
      camera.updateProjectionMatrix();

      // Update for next frame
      initialPinchDistance = currentDistance;
    }
  }
}

function onTouchEnd(event) {
  if (event.touches.length < 2) {
    initialPinchDistance = 0;
    lastTouches = [];
  }
}

function animate() {
  requestAnimationFrame(animate);

  // Update text sprites to face camera (billboard effect)
  textSprites.forEach((sprite) => {
    // Get world position of sprite
    const worldPos = new THREE.Vector3();
    sprite.getWorldPosition(worldPos);

    // Make sprite look at camera
    sprite.lookAt(camera.position);
  });

  // Update label sprites to face camera
  labelSprites.forEach((sprite) => {
    sprite.lookAt(camera.position);
  });

  // Render with anaglyph effect if enabled, otherwise normal render
  if (isAnaglyphMode) {
    anaglyphEffect.render(scene, camera);
  } else {
    renderer.render(scene, camera);
  }
}

// Toggle Anaglyph 3D mode
function toggleAnaglyph() {
  isAnaglyphMode = !isAnaglyphMode;

  if (isAnaglyphMode) {
    // Switch to perspective camera
    // Copy position and rotation from current camera
    perspectiveCamera.position.copy(camera.position);
    perspectiveCamera.rotation.copy(camera.rotation);
    camera = perspectiveCamera;
  } else {
    // Switch back to orthographic camera
    // Copy position and rotation from current camera
    orthographicCamera.position.copy(camera.position);
    orthographicCamera.rotation.copy(camera.rotation);
    camera = orthographicCamera;
  }

  return isAnaglyphMode;
}

// Toggle text formatting functions
function toggleBold() {
  if (!selectionStart) return false;

  // Get the first selected cell to check current state
  const firstKey = `${selectionStart.x},${selectionStart.y},${selectionStart.z}`;
  const currentState = cellTextBold[firstKey] || false;
  const newState = !currentState;

  // Apply to all selected cells
  const minX = Math.min(selectionStart.x, selectionEnd.x);
  const maxX = Math.max(selectionStart.x, selectionEnd.x);
  const minY = Math.min(selectionStart.y, selectionEnd.y);
  const maxY = Math.max(selectionStart.y, selectionEnd.y);
  const minZ = Math.min(selectionStart.z, selectionEnd.z);
  const maxZ = Math.max(selectionStart.z, selectionEnd.z);

  for (let x = minX; x <= maxX; x++) {
    for (let y = minY; y <= maxY; y++) {
      for (let z = minZ; z <= maxZ; z++) {
        const key = `${x},${y},${z}`;
        cellTextBold[key] = newState;

        // Update the text sprite if cell has text
        if (cellData[key]) {
          updateCellText(x, y, z, cellData[key]);
        }
      }
    }
  }

  return newState;
}

function toggleItalic() {
  if (!selectionStart) return false;

  // Get the first selected cell to check current state
  const firstKey = `${selectionStart.x},${selectionStart.y},${selectionStart.z}`;
  const currentState = cellTextItalic[firstKey] || false;
  const newState = !currentState;

  // Apply to all selected cells
  const minX = Math.min(selectionStart.x, selectionEnd.x);
  const maxX = Math.max(selectionStart.x, selectionEnd.x);
  const minY = Math.min(selectionStart.y, selectionEnd.y);
  const maxY = Math.max(selectionStart.y, selectionEnd.y);
  const minZ = Math.min(selectionStart.z, selectionEnd.z);
  const maxZ = Math.max(selectionStart.z, selectionEnd.z);

  for (let x = minX; x <= maxX; x++) {
    for (let y = minY; y <= maxY; y++) {
      for (let z = minZ; z <= maxZ; z++) {
        const key = `${x},${y},${z}`;
        cellTextItalic[key] = newState;

        // Update the text sprite if cell has text
        if (cellData[key]) {
          updateCellText(x, y, z, cellData[key]);
        }
      }
    }
  }

  return newState;
}

function toggleStrikethrough() {
  if (!selectionStart) return false;

  // Get the first selected cell to check current state
  const firstKey = `${selectionStart.x},${selectionStart.y},${selectionStart.z}`;
  const currentState = cellTextStrikethrough[firstKey] || false;
  const newState = !currentState;

  // Apply to all selected cells
  const minX = Math.min(selectionStart.x, selectionEnd.x);
  const maxX = Math.max(selectionStart.x, selectionEnd.x);
  const minY = Math.min(selectionStart.y, selectionEnd.y);
  const maxY = Math.max(selectionStart.y, selectionEnd.y);
  const minZ = Math.min(selectionStart.z, selectionEnd.z);
  const maxZ = Math.max(selectionStart.z, selectionEnd.z);

  for (let x = minX; x <= maxX; x++) {
    for (let y = minY; y <= maxY; y++) {
      for (let z = minZ; z <= maxZ; z++) {
        const key = `${x},${y},${z}`;
        cellTextStrikethrough[key] = newState;

        // Update the text sprite if cell has text
        if (cellData[key]) {
          updateCellText(x, y, z, cellData[key]);
        }
      }
    }
  }

  return newState;
}

// Set font for selected cells
function setFont(fontName) {
  if (!selectionStart) return;

  // Apply to all selected cells
  const minX = Math.min(selectionStart.x, selectionEnd.x);
  const maxX = Math.max(selectionStart.x, selectionEnd.x);
  const minY = Math.min(selectionStart.y, selectionEnd.y);
  const maxY = Math.max(selectionStart.y, selectionEnd.y);
  const minZ = Math.min(selectionStart.z, selectionEnd.z);
  const maxZ = Math.max(selectionStart.z, selectionEnd.z);

  for (let x = minX; x <= maxX; x++) {
    for (let y = minY; y <= maxY; y++) {
      for (let z = minZ; z <= maxZ; z++) {
        const key = `${x},${y},${z}`;
        cellFontFamily[key] = fontName;

        // Update the text sprite if cell has text
        if (cellData[key]) {
          updateCellText(x, y, z, cellData[key]);
        }
      }
    }
  }
}

// Expose functions to window for HTML access
window.toggleAnaglyph = toggleAnaglyph;
window.toggleBold = toggleBold;
window.toggleItalic = toggleItalic;
window.toggleStrikethrough = toggleStrikethrough;
window.setFont = setFont;

// Start the application
init();
