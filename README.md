# Service_MoveFile
A windows service to move files

It will
1. Monitor a folder
2. When triggered by a file being added to the folder, it will read an excel file (Docs\FolderDataSheet.xlsx) that specifies where files should be moved
3. The file will be renamed according to the specifications in the file, and moved
