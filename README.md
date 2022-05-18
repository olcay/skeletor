# skeletor

An automated VBA extractor from Excel workbooks. When an Excel workbook is added to the `develop` branch, its macros are extracted to a folder. A pull-request will be created automatically for the changes.

The folder name can be defined as `FileId` in the modules as a comment line `' FileId ffc35d86-8765-4c95-a099-acb8cf80a3f4`. The folder name will be `ffc35d86-8765-4c95-a099-acb8cf80a3f4.vba` in this case.

![Macro with FileId](/images/MacroWithFileId.png)

## Installation

1. Create a new repository with the [main.yml](.github/workflows/main.yml) and [extract.py](.github/workflows/extract.py) files in the `.github/workflows` directory.
1. Give the required Workflow permissions for Github Actions as in the screen-shot below.
    - Read and write permissions
    - Allow GitHub Actions to create and approve pull requests
1. Commit a new Excel workbook to the `develop` branch.
1. A new pull request will be created for the `main` branch automatically.

![Github Actions Permissions](/images/GithubActionsPermissions.png)

## Workflow Explanation

```yml
name: Extract VBA from Excel Workbooks
on:
  push:
    branches:
      - develop # Run on every commit to the develop branch.

jobs:
  run:
    name: Extraction
    runs-on: ubuntu-latest

    steps:
      - name: Checkout # Checkout the develop branch.
        uses: actions/checkout@v3
      
      - name: Install tools # Install the required tools for the Python code.
        run: python -m pip install -U oletools
        shell: sh
      
      - name: Extract # Run the Python code.
        run: python ./.github/workflows/extract.py
        shell: sh

      - name: Create Pull Request # Create a pull request from the new macros branch to the main branch.
        uses: peter-evans/create-pull-request@v4
        with:
          branch: macros
          base: main
          title: New changes in macros

```

## Python Code Explanation

The code file [extract.py](.github/workflows/extract.py) has steps below to extract macro modules from Excel workbooks.

1. Walk thorugh all the files in the root directory of the repository.
1. Remove the old extraction folders.
1. Read only the files with the defined Excel file extensions.
1. Read the module content.
1. Check the content lines for special attributes.
    - Ignore Attribute lines except VB_Name if it is needed.
    - If an id is defined then use it as a folder name.
1. If the folder does not exists, create a new one.
1. Write the content to a module file.

An explanation to the Python code that is used in this project can be found in [the sources](#sources).

## Exceptions

- If a `FileId` is not defined then the workbook name will be used.
- If there are multiple `FileId`s defined then the last one will be used.
- If a file is removed, it will not be removed from the `main` branch.

## Sources

- [How to use Git hooks to version-control your Excel VBA code](https://www.xltrail.com/blog/auto-export-vba-commit-hook)