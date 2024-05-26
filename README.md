
## MINI-EXCEL

### Overview

The "MINI-EXCEL" project is a mini-ultimate Excel sheet manager developed in C++. Inspired by the concept of a doubly linked list, the project implements a class structure where each cell in the Excel sheet is represented as a node with four links: up, down, left, and right. By leveraging these links, the system offers functionalities for managing rows, columns, and cell data within the Excel-like environment.

### Objectives

The objectives of this project are as follows:

1. **Efficient Data Management**: Provide tools and functions to efficiently manage cell data, insert rows and columns, and perform operations such as copying, cutting, and pasting data.

2. **Enhanced User Experience**: Offer a user-friendly interface with intuitive functionalities for navigating and manipulating the Excel sheet.

### Key Functions

The project includes the following key functions within the Mini Excel class:

- **InsertRowAbove()**: Inserts a row above the current cell, adjusting the links of neighboring cells accordingly.
- **InsertRowBelow()**: Inserts a row below the current cell, updating the links of adjacent cells as needed.
- **InsertColumnToRight()**: Inserts a column to the right of the current cell, managing the links of neighboring cells appropriately.
- **InsertColumnToLeft()**: Inserts a column to the left of the current cell, adjusting the links of adjacent cells as necessary.
- **InsertCellByRightShift()**: Shifts the current cell to the right and adds a new cell in its place, resulting in the insertion of a new column.
- **InsertCellByDownShift()**: Shifts the current cell down and adds a new cell in its place, leading to the insertion of a new row.
- **DeleteCellByLeftShift()**: Deletes the current cell and left shifts all cells to its right within the same row.
- **DeleteCellByUpShift()**: Deletes the current cell and shifts all cells below it within the same column one step up.
- **DeleteColumn()**: Deletes the column of the current cell, updating the links of involved cells accordingly.
- **DeleteRow()**: Deletes the row of the current cell, adjusting the links of involved cells accordingly.
- **GetRangeSum()**: Computes the sum of values within a specified range of cells.
- **GetRangeAverage()**: Calculates the average of values within a specified range of cells.
- **Iterator Class**: Implements iterators for navigating the Excel sheet in all four directions.

### Usage

1. **Compilation**: Compile the source code using a C++ compiler such as GCC or Visual Studio.
2. **Execution**: Run the compiled executable file to launch the application.
3. **Interface**: Utilize the console-based interface to interact with the Excel-like environment and perform various operations.

### File Handling

The project includes functionalities for saving and loading Excel sheets, enabling users to store and retrieve data efficiently.

### Bonus Features

Additional functionalities include resetting cell width and height when inserting characters exceeding cell dimensions.

### License

This project is licensed under the MIT License. See the `LICENSE` file for more details.

### Author

- [Muhammad Taha Saleem](https://github.com/twonum)
