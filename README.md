# 3D Excel - Isometric WebGL Spreadsheet

An isometric 3D spreadsheet application built with WebGL and Three.js, featuring orthographic camera projection and 3D cell manipulation.

## Features

- **3D Cell Grid**: 20×20×10 grid of cells with dimensions 200×200×40
- **Orthographic View**: No perspective distortion - true isometric projection
- **Interactive Editing**: Double-click cells to edit values
- **3D Navigation**: CMD+Mouse Drag to rotate and explore the 3D space
- **Billboard Text**: Cell values always face the camera for readability
- **Excel-style Labels**: Row numbers (1,2,3...) and column letters (A,B,C...)
- **Light Mode**: Clean, minimalist interface

## How to Use

1. **View Cells**: The spreadsheet starts in 2D view showing the front layer
2. **Edit Cells**: Double-click any cell to enter a value
3. **Rotate View**: Hold CMD (⌘) and drag with mouse to rotate the 3D grid
4. **Navigate Depth**: Rotate the grid to reveal cells in the Z dimension (depth layers)

## Running the Application

### Option 1: Using Vite Dev Server (Recommended)

```bash
# Install dependencies (first time only)
npm install

# Start the development server
npm run dev
```

Then open the URL shown in your terminal (usually http://localhost:5173)

### Option 2: Using Python

```bash
python3 -m http.server 8000
# Then open http://localhost:8000
```

### Option 3: Direct File Access

Simply open `index.html` in a modern web browser. The application uses ES modules and CDN-hosted Three.js, so no build process is required.

## Technical Details

- **WebGL Renderer**: Three.js r160
- **Camera**: Orthographic projection
- **Grid Size**: 20 columns × 20 rows × 10 depth layers
- **Cell Dimensions**: 200×200×40 units
- **Text Rendering**: Canvas-based sprite textures with billboard effect

## Browser Compatibility

Requires a modern browser with WebGL support (Chrome, Firefox, Safari, Edge).
