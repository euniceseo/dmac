# dmac
pipelines written for biomicro center's data management analysis core (dmac), co-op fall 2023

### figure generation code
generates a colored figure based on lab association and data association pulled from a json database

### omero scraping code
scrapes the contents of an OMERO page. code written with selenium opens up a chrome webdriver, and the code runs through all the projects, organizes the data into a parent and a child sheet, and exports the data to an excel sheet.

### sra scraping code
uses selenium to navigate parts of an sra bioproject projeect page, gather the data from each page, and organize it to be exported into an excel sheet.
