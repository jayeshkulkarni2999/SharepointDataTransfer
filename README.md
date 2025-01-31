# SharepointDataTransfer

## Overview
**SharepointDataTransfer** is a C# Console Application designed to transfer data files (**.docx, .xlsx, .pdf, .txt, etc.**) from a source SharePoint path to various destination SharePoint paths. The application intelligently categorizes and moves files based on their **entity name**, ensuring that each file is placed in the appropriate folder at its respective destination.

## Features
- Supports transfer of multiple file types (**.docx, .xlsx, .pdf, .txt, etc.**).
- Traverses nested folders in the source SharePoint location.
- Reads file metadata, including columns such as **Sr No., Document Name, Entity, etc.**.
- Categorizes files based on **Entity Name** and transfers them to their respective destination folders.
- Supports multiple destination SharePoint locations.
- Retains metadata values along with file transfer.
- **App.config** is ignored in Git to keep sensitive configurations private.

## Prerequisites
Ensure you have the following installed:
- [.NET SDK](https://dotnet.microsoft.com/download)
- [SharePoint Online Management Shell](https://www.microsoft.com/en-us/download/details.aspx?id=35588)
- [PnP PowerShell](https://pnp.github.io/powershell/)
- Git

## Setup Instructions
1. **Clone the Repository:**
   ```sh
   git clone <repository-url>
   cd SharepointDataTransfer
   ```
2. **Restore Dependencies:**
   ```sh
   dotnet restore
   ```
3. **Build the Project:**
   ```sh
   dotnet build
   ```
4. **Run the Application:**
   ```sh
   dotnet run
   ```

## Git Setup
Make sure `App.config` is ignored in `.gitignore`:
```
App.config
```

### Git Commands to Push Code
```sh
git init
git remote add origin <repo-url>
echo App.config > .gitignore
git add .
git commit -m "Initial commit"
git push -u origin main
```

## How It Works
1. The application scans the **source SharePoint location** for folders and files.
2. It reads metadata columns to identify the **Entity Name** associated with each file.
3. Based on the **Entity Name**, the application determines the correct **destination SharePoint path**.
4. It moves each file into the respective folder in the destination SharePoint.
5. Retains all metadata information along with the file transfer.

## Testing
- Ensure your SharePoint credentials and permissions are set up correctly.
- Run test transfers with sample files to validate entity-based categorization.
- Log output can be reviewed to track file transfers.

## License
This project is licensed under the MIT License.

