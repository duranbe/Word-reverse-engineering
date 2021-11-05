# Word reverse engineering 📝
Notes on reverse engineering of a Microsoft Word file

## Namespaces and tags

base namespace (**w**) is ` {http://schemas.openxmlformats.org/wordprocessingml/2006/main}`

| Tag |  Description                |
| :-------- |  :------------------------- |
| **t**  |  text |
| **p**  |  paragraph |
| **tbl**  |  table |
| **tr**  | row |
| **tc**  |  cell |
|  **sectPr**  | section |
|  **br**  | break |
|  **headerReference**  | reference to a header using rId |
|  **footerReference**  | reference to a footer using rId |

## Python 3 libs to use
- zipFile (internal lib)
- lxml (internal lib etree is renaming the namespaces)
