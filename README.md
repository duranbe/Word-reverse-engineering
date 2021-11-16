# Word reverse engineering 📝
Notes on reverse engineering of a Microsoft Word file

A Word file is in fact a compressed folder containing xml files to describe the contents, headers, footers, their relationships and images as well.

## Folder Organization

```
 - _rels/
    | .rels
 - customXML/
    | _rels/
      | item1.xml.rels
    | item1.xml
    | itemProps1.xml
 - docProps/
    | core.xml
    | custom.xml
 - word/
    | _rels/
    | media/
    | theme/
    | document.xml
    | fontTable.xml
    | footer1.xml
    | header1.xml
    | numbering.xml
    | settings.xml
    | styles.xml
 - [Content_Types].xml
```


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
|  **headerReference**  | reference to a header using **rId**, the parameter **type** can contain either 'default' for using it as default header or 'first' to use it on the first page |
|  **footerReference**  | reference to a footer using **rId**, the parameter **type** can contain either 'default' for using it as default footer or 'first' to use it on the first page |
|  **lang**  | language of text (example : en-GB) |


## XPath Expressions

Get all sections : ```//w:sectPr/...```

Get all page breaks : ```.//w:br[@w:type="page"]```

## Python 3 libs to use
- zipFile (internal lib)
- lxml (internal lib etree is renaming the namespaces)
