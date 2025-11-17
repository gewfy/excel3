# Excel³ - 3D Spreadsheet Application

A revolutionary 3D spreadsheet application that extends Excel into three dimensions. Built with Three.js and WebGL, Excel³ offers a fully interactive 3D workspace with advanced features including anaglyph 3D viewing, 4D hypercube projection, and quantum uncertainty simulation. **Made for the Swedish edition of [uoɥʇɐʞɔɐH pıdnʇS](https://stupidhackathon.se/).**

![Excel³ Banner](https://img.shields.io/badge/Excel³-3D_Spreadsheet-blue?style=for-the-badge)
![Three.js](https://img.shields.io/badge/Three.js-r160-green?style=flat-square)
![WebGL](https://img.shields.io/badge/WebGL-2.0-orange?style=flat-square)


<img width="1330" height="912" alt="image" src="https://github.com/user-attachments/assets/87b89662-2731-4bcc-bec5-7ec5c05664a7" />

## AI (but informative) slop ⬇️

### ✨ Features

#### Core Functionality

- **Dynamic 3D Grid**: Automatically sized grid based on viewport dimensions with 5 depth layers
- **Real-time Inline Editing**: Type directly into cells without modal dialogs
- **Smart Navigation**: Arrow keys adapt to viewing angle - navigate naturally from any perspective
- **Cell Selection**: Single cells, ranges, and 3D cubic regions with shift-click extension
- **AutoSum**: Intelligent summation that places results below selected ranges

#### Rich Text Formatting

- **Text Styles**: Bold, italic, and strikethrough formatting
- **Typography**: 10 font families including System, Arial, Times New Roman, Courier, and more
- **Font Sizing**: Adjustable font sizes from 20px to 300px
- **Colors**: Full color picker support for both text and cell backgrounds

#### Advanced 3D Features

- **Orthographic Camera**: True isometric view without perspective distortion (default)
- **Perspective Modes**: Switch to perspective or extreme wide-angle FOV cameras
- **Free Rotation**: Cmd/Ctrl + drag to rotate the matrix in any direction
- **Zoom & Pan**: Mouse wheel to zoom, right-click drag to pan
- **Billboard Text**: Cell values and labels always face the camera for optimal readability

#### Special Viewing Modes

- **🥽 Anaglyph 3D**: Red-cyan stereoscopic 3D mode for 3D glasses
- **🌀 4D Hypercube**: Psychedelic 4D-to-3D projection with barrel distortion
- **⚛️ Quantum Uncertainty**: Watch numeric values fluctuate until observed (clicked)
- **🔲 Hide Borders**: Toggle cell borders for a cleaner view

#### Data Persistence

- **Version Control**: Save multiple named versions of your spreadsheet
- **Auto-Save**: Quick save/load to browser localStorage
- **Version Management**: List, load, and delete saved versions
- **Full State**: Preserves all data, colors, formatting, and styles

### 🚀 Getting Started

#### Installation

```bash
# Install dependencies
npm install

# Start the development server
npm run dev
```

The application will open at `http://localhost:5173`

### 📖 User Guide

#### Basic Navigation

##### Mouse Controls

- **Click**: Select a cell
- **Click + Drag**: Select a range of cells (creates 3D cubic selection)
- **Shift + Click**: Extend selection from current cell
- **Cmd/Ctrl + Drag**: Rotate the 3D matrix
- **Right-Click + Drag**: Pan the camera view
- **Mouse Wheel**: Zoom in/out

##### Keyboard Controls

**Cell Navigation** (adapts to viewing angle):

- `↑↓←→`: Navigate cells - horizontal keys adapt based on rotation
  - Front view: Left/Right = X-axis (columns)
  - Side view: Left/Right = Z-axis (depth)
- `Alt + ↑↓`: Navigate through depth axis (swaps with columns when viewing from side)
- `Shift + Arrows`: Extend selection while navigating

**Cell Editing**:

- `Type`: Start editing selected cell
- `Enter`: Save and move to cell below
- `Escape`: Cancel editing
- `Delete/Backspace`: Clear selected cells

#### Text Formatting

Select cells, then use the toolbar:

- **B** - Bold text
- **I** - Italic text
- **S** - Strikethrough
- **A⁺** - Increase font size (+10px)
- **A⁻** - Decrease font size (-10px)
- **Font Dropdown** - Change font family
- **🎨 Color Pickers** - Set background and text colors

#### Special Features

##### AutoSum (Σ)

1. Select a range of cells with numeric values
2. Click the Σ button in toolbar
3. Sum appears in the cell immediately below your selection
4. Original selection remains active

##### Version Management

- **Save**: Opens version save dialog - enter a name and save
- **Load**: View list of saved versions, click to load or delete
- Each version stores complete state including all formatting

##### Special View Modes

**Anaglyph 3D (3D Glasses Icon)**:

- Enables red-cyan stereoscopic 3D
- Put on red-cyan 3D glasses for depth perception
- Automatically switches to perspective camera

**4D Hypercube (4D Icon)**:

- Projects 4D coordinates into 3D space
- Continuous rotation through 4th dimension
- Barrel distortion creates fisheye effect
- Switches to ultra-wide FOV camera

**Quantum Uncertainty (⚛️ Icon)**:

- Numeric values fluctuate ±10%
- Click a cell to "observe" and collapse the wave function
- Values freeze once observed
- Great for demonstrating quantum superposition

**Hide Borders (Grid Icon)**:

- Toggles cell border visibility
- Cleaner view for presentations

#### Tips & Tricks

1. **Navigate at Any Angle**: Rotate the matrix 90° to the side, and left/right arrows will naturally navigate through depth layers
2. **3D Selection**: Drag across cells while rotated to select 3D regions
3. **Persistent Data**: All your work auto-saves to localStorage when you click save
4. **Mobile Support**: Works on touch devices with pinch-to-zoom
5. **Performance**: Grid dynamically adjusts to viewport size for optimal performance

## 🤝 Contributing

This is a demonstration project showcasing 3D web technologies. Feel free to fork and extend!

---

Built with ❤️ and curiosity about dimensional spreadsheets.
