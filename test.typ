#set page(width: 20cm, height:auto)
#set heading(numbering: "1.")

= Fun with typst

The typst program is pretty neat! 

- Extremely fast incremental recompile
- Easier syntax
- Markup includes a well-designed scripting language
  (vs. LaTeX's macro system; akin to the difference between
  using macros in the C preprocessor vs. functions)

Circles #box[#circle(radius: 4pt)] are easy to embed, 
along with equations like $E = m c^2$ and $a^2+b^2=c^2$ 
and $sum_(k=0)^infinity 1/(k^2)=pi^2 / 6$. Or on its own line:

#let bf(x) = $upright(bold(#x))$
#set math.mat(delim: "[")

$ bf(T)(theta) = mat(cos theta, -sin theta; sin theta, cos theta) $

- `typst` program: https://github.com/typst/typst/releases/tag/22-03-21-2
- vscode language files for syntax highlighting at https://github.com/typst/typst/tree/main/tools/support

  - Download the vscode support files

    ```
    mkdir typst-vscode
    cd typst-vscode
    git clone --filter=blob:none --sparse https://github.com/typst/typst .
    git sparse-checkout add tools/support
    ```
  - In VSCode:
      - Shift-Ctrl-P
      - select "Developer: Install Extension from Location"
      - select the tools/support directory you got from git
- Run `typst --watch MYFILE.typ` in a terminal in VS Code
- Install #link("https://marketplace.visualstudio.com/items?itemName=tomoki1207.pdf")[vscode-pdf extension]
    for viewing pdf
    
/*
 * Here is a sample comment
 *
 * typst theses:
 * https://www.user.tu-berlin.de/laurmaedje/programmable-markup-language-for-typesetting.pdf
 * https://www.user.tu-berlin.de/mhaug/fast-typesetting-incremental-compilation.pdf
 */

#heading(level: 2)[Listing 1 - Changing Font in Excel with TypeScript]

#heading(level: 3)[2.1 Example Subsection]

This is a sample subsection under 2.1.
= Kod
#heading(level: 2)[Detailed Example]

This is a detailed example under subsection 2.1.1.

#box(
  inset: 8pt,
  fill: rgb("#f0f0f0"),
  radius: 4pt
)[
  #figure(
    caption: [Listing 1: Changing Font in Excel with TypeScript]
  )[
    #raw(
      "// TypeScript code to change the font of all cells to Calibri in an Excel worksheet\nimport Excel from 'exceljs';\n\nasync function changeFontToCalibri(workbookPath: string): Promise<void> {\n  try {\n    // Load the workbook\n    const workbook = new Excel.Workbook();\n    await workbook.xlsx.readFile(workbookPath);\n\n    // Iterate through all the worksheets in the workbook\n    workbook.eachSheet((worksheet) => {\n      // Iterate through all rows in the worksheet\n      worksheet.eachRow((row) => {\n        // Iterate through all cells in each row\n        row.eachCell((cell) => {\n          // Set the font of the cell to Calibri\n          cell.font = {\n            name: 'Calibri',\n          };\n        });\n      });\n    });\n\n    // Save the workbook after modifications\n    await workbook.xlsx.writeFile(workbookPath);\n    console.log('Font changed to Calibri for all cells successfully.');\n  } catch (error) {\n    console.error('An error occurred:', error);\n  }\n}\n\n// Usage\nconst workbookPath = 'path/to/your/excel/file.xlsx';\nchangeFontToCalibri(workbookPath);",
      block: true,
      lang: "typescript"
    )
  ]
]



That's how easy it is to add a code snippet into a `typst` document!
