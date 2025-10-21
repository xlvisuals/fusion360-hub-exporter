# fusion360-hub-exporter
Export all Fusion 360 projects and designs from the active hub to disk

## Capabilities
- Exports every project as .f3d, .stp, .stl, and .igs (configurable)
- Exports a screenshot of the project (optional)
- Exports component bodies as .stl (optional)
- Exports component sketches as .dxf (optional)

## How to use:
1. Download and extract
2. Configure export options in Fusion360HubExporter.py
3. Open Fusion 360
4. Click on "Utilities" > "Add-ins" > "Scripts and Add-ins"
5. Click the + button and select the folder with this script
6. Confirm popup that this might take a while
7. Select target folder
8. Wait while fusion opens every project and design, exports it, then closes it.

## Acknowledgements:
This program is based on the work of Justin Nesselrotte - with gratitude.
https://github.com/Jnesselr/fusion-360-total-exporter
