# e2s
Small app to convert Nikon's e-max files (*.tch extension) into dxf files or directly into Solidworks sketches (including 3D points)
Requires pywin32 if using Solidworks API (builds 216 and older work fine, builds after 216 didn't work with Solidworks API)
Requires dxfwrite module to export data to dxf file.
3D points cloud can be exported to dxf file, changing microscope into kind of 3D scanner (with not very accurate z axis, though).
