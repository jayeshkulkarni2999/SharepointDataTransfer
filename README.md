# SharepointDataTransfer

## Overview
**SharepointDataTransfer** is a C# Console Application that utilizes **SharePoint APIs** to transfer data files (**.docx, .xlsx, .pdf, .txt, etc.**) from a source SharePoint location to multiple destination SharePoint locations. The application intelligently categorizes and moves files based on their **entity name**, ensuring that each file is placed in the correct destination folder while retaining metadata.

## Features
- Utilizes **SharePoint APIs** to interact with SharePoint Online.
- Supports transfer of multiple file types (**.docx, .xlsx, .pdf, .txt, etc.**).
- Traverses nested folders in the source SharePoint location.
- Reads file metadata, including columns such as **Sr No., Document Name, Entity, etc.**.
- Categorizes files based on **Entity Name** and transfers them to their respective SharePoint sites/folders.
- Supports multiple destination SharePoint locations.
- Retains metadata values along with file transfer.
- **App.config** is ignored in Git to keep sensitive configurations private.

## Prerequisites
Ensure you have the following installed:
- [.NET SDK](https://dotnet.microsoft.com/download)
- [SharePoint Online Management Shell](https://www.microsoft.com/en-us/download/details.aspx?id=35588)
- [PnP PowerShell](https://pnp.github.io/powershell/)
- [SharePoint Client-Side Object Model (CSOM)](https://www.nuget.org/packages/Microsoft.SharePointOnline.CSOM)
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
1. The application connects to the **source SharePoint site** using **SharePoint APIs**.
2. It scans folders and subfolders, retrieving file metadata.
3. The **Entity Name** from metadata determines the correct **destination SharePoint site/folder**.
4. The application uploads the file to the appropriate **destination SharePoint** location.
5. Metadata is retained and mapped appropriately during transfer.

## Configuration
- The **App.config** file stores SharePoint site URLs, authentication details, and folder mappings.
- Ensure that necessary permissions are granted for the SharePoint API.

## Testing
- Ensure SharePoint API credentials are configured correctly.
- Run test transfers with sample files to validate entity-based categorization.
- Log output can be reviewed to track file transfers and errors.

## License
This project is licensed under the MIT License.

