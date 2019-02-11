### Overview

Provides facility for abstracting away the cell model of PhpSpreadsheet,
and utilizing tablular concepts such as tables and columns.
This package essentially saves the programmer from having to constantly 
remember what cell represents what table entry.

### Dependencies
    PHP 7.2
    PhpSpreadSheet

### Installing
Just run `composer install` to get the dependencies.

### Utilizing
Create the table model utilizing the Sheet* classes and then utilize the 
SpreadsheetTableFacade to render/map conceptual table model to a
PHPSpreadsheet worksheet.

For examples look at the ./tests folder in the project root directory.

### Reading the code
The Facade classes responsibility is to map the conceptual model to a
PHPSpreadsheet, and thus contains all the logic related to PHPSpreadsheet.

The conceptual model is based on anchor points, and calculating the elements
appropriate anchor point through `resolveAddress()`. It is recommended to
start with the AnchorableEntity class.
