# Create-PDF-via-Python             
1.Not yet validated                        
Not yet validated or confirmed by the Takeda team. However, based on the conversion results, the output appears to comply with the FDA Portable Document Format (PDF) Specifications and other applicable guidelines. Additional compliance verification is currently being pursued.

2.Core process                              
Adobe Acrobat Pro:
Word → [Word COM ExportAsFixedFormat] → PDF → Acrobat Pro Modify properties

Python:
Word → [Word COM ExportAsFixedFormat] → PDF → Python pikepdf Modify properties

3.Must executed in J drive                         
Must be executed from the J: drive, as the Python environment is already configured there. The .vbs and .py files can be copied to any folder on the J: drive and then executed. To run the tool locally, the Python environment must be configured manually.

4.Run                      
- Double-click the .vbs files in the order of Steps 1 and 2, then wait for several seconds for the process to complete.
- The Fast Web View optimization must be completed manually by using the Save As function.

Note: 2.vbs supports batch checking of all PDF files in a folder (excluding subfolders). This can be helpful for quickly validating PDF files when an ADaM eCRT package contains a large number of PDFs.

------------------------------------------------
2026-08-04:              
The tool may also be supported in the SCE folder once Satori provides Python environment support.
