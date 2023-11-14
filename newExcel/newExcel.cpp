
#include <iostream>
#include <vector>
#include<Windows.h>
#include <string>
#include <fstream>
#include <sstream>
#include <limits.h>
using namespace std;

//converting any data type to string
template<typename T>
string to_string(T value)
{
	ostringstream os;
	os << value;
	return os.str();
}
template <typename T>
class Cell {
public:
	T value;
	bool isActive;
	int rowNo, colNo;
	Cell<T>* leftCell;
	Cell<T>* rightCell;
	Cell<T>* upperCell;
	Cell<T>* lowerCell;
	string color;
	Cell(T data) {
		this->value = data;
		isActive = true;
		upperCell = nullptr;
		lowerCell = nullptr;
		leftCell = nullptr;
		rightCell = nullptr;
	}
	Cell() {
		value = "0";
		isActive = false;
		color = "\33[37m";
		upperCell = nullptr;
		lowerCell = nullptr;
		leftCell = nullptr;
		rightCell = nullptr;
	}
};
template <typename T>
class Spreadsheet {
public:
	int numRows;
	int numCols;
	Cell<T>* headCell;
	Cell<T>* currCell;
	Cell<T>* prevRowCell;
	Cell<T>* currRowCell;
	Cell<T>* originalHeadCell;
	Spreadsheet() {
		this->numRows = 5;
		this->numCols = 5;
		headCell = new Cell<T>();
		currCell = headCell;
		createGrid();
	}
	Spreadsheet(int rows, int cols) {
		this->numRows = rows;
		this->numCols = cols;
		headCell = new Cell<T>();
		currCell = headCell;
		createGrid();
	}
	void createGrid();
	void displayGrid();
	void insertRowAbove();
	void insertRowBelow();
	void insertColumnToRight();
	void insertColumnToLeft();
	void insertCellByRightShift();
	void insertCellByDownShift();
	void deleteCellByLeftShift();
	void deleteCellByUpShift();
	void deleteColumn();
	void deleteRow();
	void clearColumn();
	void clearRow();
	//void getRangeSum(int startRow, int startCol, int endRow, int endCol);
	//void getRangeAverage(int startRow, int startCol, int endRow, int endCol);
	//void getRangeCount(int startRow, int startCol, int endRow, int endCol);
	//void getRangeMin(int startRow, int startCol, int endRow, int endCol);
	//void getRangeMax(int startRow, int startCol, int endRow, int endCol);
	void save();
	void load();
	void resetCellWidth();
	void resetCellHeight();
	/*void makeColumnAtEndOrBottom(Cell<T>* head) {
		Cell<T>* newCell = new Cell<T>();
		Cell<T>* tempDown = head->lowerCell;
		head->rightCell = newCell;
		newCell->leftCell = head;
		Cell<T>* newRowPointer = newCell;
		for (int i = 1; i < numRows; i++) {
			Cell<T>* newCells = new Cell<T>();
			newRowPointer->lowerCell = newCells;
			newCells->upperCell = newRowPointer;
			tempDown->rightCell = newCells;
			newCells->leftCell = tempDown;
			newRowPointer = newCells;
			tempDown = tempDown->lowerCell;
		}
	}*/
	void makeRightCol() {
		Cell<T>* temp = headCell;
		while (temp->rightCell != nullptr)
		{
			temp = temp->rightCell;
		}
		Cell<T>* newCell = new Cell<T>();
		Cell<T>* tempDown = temp->lowerCell;
		temp->rightCell = newCell;
		newCell->leftCell = temp;
		Cell<T>* newRowPointer = newCell;
		for (int i = 1; i < numRows; i++) {
			Cell<T>* newCells = new Cell<T>();
			newRowPointer->lowerCell = newCells;
			newCells->upperCell = newRowPointer;
			tempDown->rightCell = newCells;
			newCells->leftCell = tempDown;
			newRowPointer = newCells;
			tempDown = tempDown->lowerCell;
		}
	}
	void makeNewRowAtEnd() {
		Cell<T>* temp = headCell;
		while (temp->lowerCell != nullptr)
		{
			temp = temp->lowerCell;
		}
		Cell<T>* newCell = new Cell<T>();
		Cell<T>* tempRight = temp->rightCell;
		temp->lowerCell = newCell;
		newCell->upperCell = temp;
		Cell<T>* newRowPointer = newCell;
		for (int i = 1; i < numCols; i++) {
			Cell<T>* newCells = new Cell<T>();
			newRowPointer->rightCell = newCells;
			newCells->leftCell = newRowPointer;
			tempRight->lowerCell = newCells;
			newCells->upperCell = tempRight;
			tempRight = tempRight->rightCell;
			newRowPointer = newCells;
		}

	}
	void makeLeftNewCol() {
		Cell<T>* temp = headCell;
		Cell<T>* newCell = new Cell<T>();
		Cell<T>* tempDown = temp->lowerCell;
		temp->leftCell = newCell;
		newCell->rightCell = temp;
		headCell = newCell;
		for (int i = 1; i < numRows; i++) {
			Cell<T>* newCells = new Cell<T>();
			newCell->lowerCell = newCells;
			newCells->upperCell = newCell;
			tempDown->leftCell = newCells;
			newCells->rightCell = tempDown;
			newCell = newCells;
			tempDown = tempDown->lowerCell;
		}
	}
	void newRowAtTop() {
		Cell<T>* newRowHead = new Cell<T>();
		Cell<T>* temp = headCell->rightCell;
		headCell->upperCell = newRowHead;
		newRowHead->lowerCell = headCell;
		Cell<T>* curr = newRowHead;
		for (int i = 1; i < numCols; i++) {
			Cell<T>* newCell = new Cell<T>();
			newCell->leftCell = curr;
			curr->rightCell = newCell;
			temp->upperCell = newCell;
			newCell->lowerCell = temp;
			temp = temp->rightCell;
			curr = newCell;
		}
		headCell = newRowHead;
	}
};
template <typename T>
void Spreadsheet<T>::createGrid() {
	Cell<T>* tempRowCell = headCell;
	for (int i = 1; i < numRows; i++)
	{
		Cell<T>* newCell = new Cell<T>();
		tempRowCell->lowerCell = newCell;
		newCell->upperCell = tempRowCell;
		tempRowCell = tempRowCell->lowerCell;
	}
	tempRowCell = headCell;
	for (int i = 1; i < numCols; i++)
	{
		Cell<T>* newCell = new Cell<T>();
		tempRowCell->rightCell = newCell;
		newCell->leftCell = tempRowCell;
		tempRowCell = tempRowCell->rightCell;
	}
	tempRowCell = headCell;

	Cell<T>* last = headCell;
	Cell<T>* currRight = headCell->rightCell;
	for (int i = 1; i < numCols; i++) {
		Cell<T>* preDown = last->lowerCell;
		Cell<T>* curr = currRight;
		for (int j = 1; j < numRows; j++) {
			Cell<T>* newCell = new Cell<T>();
			preDown->rightCell = newCell;
			newCell->leftCell = preDown;
			preDown = preDown->lowerCell;
			curr->lowerCell = newCell;
			newCell->upperCell = curr;
			curr = newCell;
		}
		last = currRight;
		currRight = currRight->rightCell;
	}
}

template <typename T>
void Spreadsheet<T>::displayGrid() {
	Cell<T>* temp = headCell;
	Cell<T>* tempRow;

	// Display column labels
	cout << "                                                   ";
	for (int i = 0; i < numCols; i++) {
		char colLabel = 'A' + i;
		cout << colLabel << "    ";
	}
	cout << endl;

	// Display cell values
	for (int i = 0; i < numRows; i++) {
		cout << "                                          ";
		cout << i + 1 << " | -> ";

		tempRow = temp;

		for (int j = 0; j < numCols; j++) {
			cout << " [" << tempRow->color << tempRow->value << "] ";
			tempRow = tempRow->rightCell;
		}

		cout << endl;
		temp = temp->lowerCell;
	}
}


template <typename T>
void Spreadsheet<T>::insertRowAbove() {
	Cell<T>* temp = headCell;
	while (temp->leftCell != nullptr)
	{
		temp = temp->leftCell;
	}
	//if there is row above the current row
	if (temp->upperCell) {
		Cell<T>* tempUp = temp->upperCell;
		Cell<T>* tempRight = temp->rightCell;
		Cell<T>* currRight = temp->rightCell;
		Cell<T>* newCell = new Cell<T>();
		tempUp->lowerCell = newCell;
		newCell->upperCell = tempUp;
		temp->upperCell = newCell;
		newCell->lowerCell = temp;

		Cell<T>* newRowPointer = newCell;

		for (int i = 1; i < numCols; i++) {
			Cell<T>* newCells = new Cell<T>();
			newRowPointer->rightCell = newCells;
			newCells->leftCell = newRowPointer;
			tempRight->lowerCell = newCells;
			newCells->upperCell = tempRight;
			currRight->upperCell = newCells;
			newCells->lowerCell = currRight;
			newRowPointer = newCells;
			tempRight = tempRight->rightCell;
			currRight = currRight->rightCell;
		}
	}
	//if there is no row above the current row
	else {
		/*Cell<T>* newCell = new Cell<T>();
		Cell<T>* tempRight = headCell->rightCell;
		headCell->upperCell = newCell;
		newCell->lowerCell = headCell;
		Cell<T>* newRowPointer = newCell;
		for (int i = 1; i < numCols; i++) {
			Cell<T>* newCells = new Cell<T>();
			newRowPointer->rightCell = newCells;
			newCells->leftCell = newRowPointer;
			tempRight->upperCell = newCells;
			newCells->lowerCell = tempRight;
			tempRight = tempRight->rightCell;
			newRowPointer = newCells;
		}
		headCell = newCell;*/
		newRowAtTop();
	}
	numRows++;
}

template <typename T>
void Spreadsheet<T>::insertRowBelow() {
	Cell<T>* temp = headCell;
	while (temp->leftCell != nullptr)
	{
		temp = temp->leftCell;
	}
	//if there is row below the current row
	if (temp->lowerCell) {
		Cell<T>* tempDown = temp->lowerCell;
		Cell<T>* tempRight = temp->rightCell;
		Cell<T>* currRight = temp->rightCell;
		Cell<T>* newCell = new Cell<T>();
		tempDown->upperCell = newCell;
		newCell->lowerCell = tempDown;
		temp->lowerCell = newCell;
		newCell->upperCell = temp;

		Cell<T>* newRowPointer = newCell;

		for (int i = 1; i < numCols; i++) {
			Cell<T>* newCells = new Cell<T>();
			newRowPointer->rightCell = newCells;
			newCells->leftCell = newRowPointer;
			tempRight->upperCell = newCells;
			newCells->lowerCell = tempRight;
			currRight->lowerCell = newCells;
			newCells->upperCell = currRight;
			newRowPointer = newCells;
			tempRight = tempRight->rightCell;
			currRight = currRight->rightCell;
		}
	}
	//if there is no row below the current row
	else {
		/*Cell<T>* newCell = new Cell<T>();
		Cell<T>* tempRight = headCell->rightCell;
		headCell->lowerCell = newCell;
		newCell->upperCell = headCell;
		Cell<T>* newRowPointer = newCell;
		for (int i = 1; i < numCols; i++) {
			Cell<T>* newCells = new Cell<T>();
			newRowPointer->rightCell = newCells;
			newCells->leftCell = newRowPointer;
			tempRight->lowerCell = newCells;
			newCells->upperCell = tempRight;
			tempRight = tempRight->rightCell;
			newRowPointer = newCells;
		}*/
		makeNewRowAtEnd();
	}
	numRows++;
}
template <typename T>
void Spreadsheet<T>::insertColumnToRight() {
	Cell<T>* firstHead = headCell;
	while (firstHead->upperCell != nullptr)
	{
		firstHead = firstHead->upperCell;
	}
	if (firstHead->rightCell) {
		Cell<T>* tempRight = firstHead->rightCell;
		Cell<T>* tempDown = firstHead->lowerCell;
		Cell<T>* currDown = firstHead->lowerCell;
		Cell<T>* newCell = new Cell<T>();
		newCell->leftCell = firstHead;
		firstHead->rightCell = newCell;
		newCell->rightCell = tempRight;
		tempRight->leftCell = newCell;

		for (int i = 1; i < numRows; i++) {
			Cell<T>* newCells = new Cell<T>();
			newCell->lowerCell = newCells;
			newCells->upperCell = newCell;
			tempDown->leftCell = newCells;
			newCells->rightCell = tempDown;
			currDown->rightCell = newCells;
			newCells->leftCell = currDown;
			newCell = newCells;
			tempDown = tempDown->lowerCell;
			currDown = currDown->lowerCell;
		}
	}
	else {
		/*Cell<T>* tempDown = firstHead->lowerCell;
		Cell<T>* currDown = firstHead->lowerCell;
		Cell<T>* newCell = new Cell<T>();
		newCell->leftCell = firstHead;
		firstHead->rightCell = newCell;
		for (int i = 1; i < numRows; i++) {
			Cell<T>* newCells = new Cell<T>();
			newCell->lowerCell = newCells;
			newCells->upperCell = newCell;
			tempDown->leftCell = newCells;
			newCells->rightCell = tempDown;
			currDown->rightCell = newCells;
			newCells->leftCell = currDown;
			newCell = newCells;
			tempDown = tempDown->lowerCell;
			currDown = currDown->lowerCell;
		}*/
		makeRightCol();
	}
	numCols++;
}
template <typename T>
void Spreadsheet<T>::insertColumnToLeft() {
	Cell<T>* firstHead = headCell;
	while (firstHead->upperCell != nullptr)
	{
		firstHead = firstHead->upperCell;
	}
	//if there is column to the left of current column
	if (firstHead->leftCell) {
		Cell<T>* tempLeft = firstHead->leftCell;
		Cell<T>* tempDown = firstHead->lowerCell;
		Cell<T>* currDown = firstHead->lowerCell;
		Cell<T>* newCell = new Cell<T>();
		newCell->rightCell = firstHead;
		firstHead->leftCell = newCell;
		newCell->leftCell = tempLeft;
		tempLeft->rightCell = newCell;

		for (int i = 1; i < numRows; i++) {
			Cell<T>* newCells = new Cell<T>();
			newCell->lowerCell = newCells;
			newCells->upperCell = newCell;
			tempDown->rightCell = newCells;
			newCells->leftCell = tempDown;
			currDown->leftCell = newCells;
			newCells->rightCell = currDown;
			newCell = newCells;
			tempDown = tempDown->lowerCell;
			currDown = currDown->lowerCell;
		}
	}
	//if there is no column to the left of current column
	else {
		makeLeftNewCol();
	}
	numCols++;
}
template <typename T>
void Spreadsheet<T>::insertCellByRightShift() {
	Cell<T>* temp = currCell;
	while (temp->rightCell != nullptr)
	{
		temp = temp->rightCell;
	}
	while (temp->upperCell != nullptr)
	{
		temp = temp->upperCell;
	}
	makeRightCol();
	//right shift all the cells
	Cell<T>* curr = currCell;
	temp = curr->leftCell;
	Cell<T>* newCell = new Cell<T>();
	if (temp) {
		temp->rightCell = newCell;
		newCell->leftCell = temp;
	}
	newCell->rightCell = curr;
	curr->leftCell = newCell;
	temp = newCell;
	//if it is first row
	if (curr->upperCell == nullptr) {
		Cell<T>* tempDown = curr->lowerCell;
		if (currCell == headCell)
			headCell = newCell;
		while (tempDown != nullptr)
		{
			temp->lowerCell = tempDown;
			tempDown->upperCell = temp;
			temp = temp->rightCell;
			tempDown = tempDown->rightCell;
		}
	}
	else if (curr->lowerCell == nullptr) {
		Cell<T>* tempUp = curr->upperCell;
		while (tempUp != nullptr)
		{
			temp->upperCell = tempUp;
			tempUp->lowerCell = temp;
			temp = temp->rightCell;
			tempUp = tempUp->rightCell;
		}
	}
	else {
		Cell<T>* tempUp = curr->upperCell;
		Cell<T>* tempDown = curr->lowerCell;
		while (tempUp != nullptr)
		{
			temp->upperCell = tempUp;
			tempUp->lowerCell = temp;
			tempDown->upperCell = temp;
			temp->lowerCell = tempDown;
			temp = temp->rightCell;
			tempUp = tempUp->rightCell;
			tempDown = tempDown->rightCell;
		}
	}
	temp->leftCell->rightCell = nullptr;
	numCols++;
}
template <typename T>
void Spreadsheet<T>::insertCellByDownShift() {
	Cell<T>* temp = currCell;
	while (temp->lowerCell != nullptr)
	{
		temp = temp->lowerCell;
	}
	while (temp->leftCell != nullptr)
	{
		temp = temp->leftCell;
	}
	makeNewRowAtEnd();
	//down shift all the cells
	Cell<T>* curr = currCell;
	temp = curr->upperCell;
	Cell<T>* newCell = new Cell<T>();
	if (temp) {
		temp->lowerCell = newCell;
		newCell->upperCell = temp;
	}
	newCell->lowerCell = curr;
	curr->upperCell = newCell;
	temp = newCell;
	//if it is first column
	if (!curr->leftCell) {
		Cell<T>* tempRight = curr->rightCell;
		if (currCell == headCell) {

			headCell = newCell;
		}
		while (tempRight != nullptr)
		{
			temp->rightCell = tempRight;
			tempRight->leftCell = temp;
			temp = temp->lowerCell;
			tempRight = tempRight->lowerCell;
		}
	}
	else if (!curr->rightCell) {
		Cell<T>* tempLeft = curr->leftCell;
		while (tempLeft != nullptr)
		{
			temp->leftCell = tempLeft;
			tempLeft->rightCell = temp;
			temp = temp->lowerCell;
			tempLeft = tempLeft->lowerCell;
		}
	}
	else {
		Cell<T>* tempRight = curr->rightCell;
		Cell<T>* tempLeft = curr->leftCell;
		while (tempRight != nullptr)
		{
			temp->rightCell = tempRight;
			tempRight->leftCell = temp;
			tempLeft->rightCell = temp;
			temp->leftCell = tempLeft;
			temp = temp->lowerCell;
			tempRight = tempRight->lowerCell;
			tempLeft = tempLeft->lowerCell;
		}
	}
	temp->upperCell->lowerCell = nullptr;
	numRows++;
}
template <typename T>
void Spreadsheet<T>::deleteCellByLeftShift() {
	Cell<T>* active = currCell;
	Cell<T>* temp = active;
	while (temp->rightCell != nullptr)
	{
		temp = temp->rightCell;
	}
	Cell<T>* newCell = new Cell<T>();
	temp->rightCell = newCell;
	newCell->leftCell = temp;
	//delete the active cell
	temp = active->leftCell;
	if (temp) {
		temp->rightCell = active->rightCell;
		active->rightCell->leftCell = temp;
	}
	else {
		active->rightCell->leftCell = nullptr;
	}
	//left shift all the cells
	if (active->upperCell == nullptr) {
		Cell<T>* tempDown = active->lowerCell;
		temp = active->rightCell;
		while (tempDown != nullptr)
		{
			tempDown->upperCell = temp;
			temp->lowerCell = tempDown;
			temp = temp->rightCell;
			tempDown = tempDown->rightCell;
		}
		if (currCell == headCell) {
			headCell = active->rightCell;
		}
	}
	else if (active->lowerCell == nullptr) {
		Cell<T>* tempUp = active->upperCell;
		temp = active->rightCell;
		while (tempUp != nullptr)
		{
			tempUp->lowerCell = temp;
			temp->upperCell = tempUp;
			temp = temp->rightCell;
			tempUp = tempUp->rightCell;
		}
	}
	else {
		Cell<T>* tempUp = active->upperCell;
		Cell<T>* tempDown = active->lowerCell;
		temp = active->rightCell;
		while (tempUp != nullptr)
		{
			tempUp->lowerCell = temp;
			temp->upperCell = tempUp;
			tempDown->upperCell = temp;
			temp->lowerCell = tempDown;
			temp = temp->rightCell;
			tempUp = tempUp->rightCell;
			tempDown = tempDown->rightCell;
		}
	}
	currCell = active->rightCell;
}
template <typename T>
void Spreadsheet<T>::deleteCellByUpShift() {
	Cell<T>* active = currCell;
	Cell<T>* temp = active;
	while (temp->lowerCell != nullptr)
	{
		temp = temp->lowerCell;
	}
	Cell<T>* newCell = new Cell<T>();
	temp->lowerCell = newCell;
	newCell->upperCell = temp;
	//delete the active cell
	temp = active->upperCell;
	if (temp) {
		temp->lowerCell = active->lowerCell;
		active->lowerCell->upperCell = temp;
	}
	else {
		active->lowerCell->upperCell = nullptr;
	}
	//up shift all the cells
	if (active->leftCell == nullptr) {
		Cell<T>* tempRight = active->rightCell;
		temp = active->lowerCell;
		while (tempRight != nullptr)
		{
			tempRight->leftCell = temp;
			temp->rightCell = tempRight;
			temp = temp->lowerCell;
			tempRight = tempRight->lowerCell;
		}
		if (currCell == headCell) {
			headCell = active->lowerCell;
		}
	}
	else if (active->rightCell == nullptr) {
		Cell<T>* tempLeft = active->leftCell;
		temp = active->lowerCell;
		while (tempLeft != nullptr)
		{
			tempLeft->rightCell = temp;
			temp->leftCell = tempLeft;
			temp = temp->lowerCell;
			tempLeft = tempLeft->lowerCell;
		}
	}
	else {
		Cell<T>* tempRight = active->rightCell;
		Cell<T>* tempLeft = active->leftCell;
		temp = active->lowerCell;
		while (tempRight != nullptr)
		{
			tempRight->leftCell = temp;
			temp->rightCell = tempRight;
			tempLeft->rightCell = temp;
			temp->leftCell = tempLeft;
			temp = temp->lowerCell;
			tempRight = tempRight->lowerCell;
			tempLeft = tempLeft->lowerCell;
		}
	}
	currCell = active->lowerCell;
}
template <typename T>
void Spreadsheet<T>::deleteColumn() {
	Cell<T>* active = currCell;
	Cell<T>* temp = active;
	while (temp->upperCell != nullptr)
	{
		temp = temp->upperCell;
	}
	if (!temp->leftCell) {
		headCell = temp->rightCell;
		Cell<T>* tempRight = temp->rightCell;
		while (tempRight != nullptr)
		{
			tempRight->leftCell = nullptr;
			tempRight = tempRight->lowerCell;
		}
		headCell = temp->rightCell;
	}
	else if (!temp->rightCell) {
		Cell<T>* tempLeft = temp->leftCell;
		while (tempLeft != nullptr)
		{
			tempLeft->rightCell = nullptr;
			tempLeft = tempLeft->leftCell;
		}
		currCell = active->leftCell;
	}
	else {
		Cell<T>* tempRight = temp->rightCell;
		Cell<T>* tempLeft = temp->leftCell;
		while (tempLeft != nullptr)
		{
			tempLeft->rightCell = tempRight;
			tempRight->leftCell = tempLeft;
			tempLeft = tempLeft->lowerCell;
			tempRight = tempRight->lowerCell;
		}
		currCell = active->leftCell;
	}
	numCols--;
}
template <typename T>
void Spreadsheet<T>::deleteRow() {
	Cell<T>* active = currCell;
	Cell<T>* temp = active;
	while (temp->leftCell != nullptr)
	{
		temp = temp->leftCell;
	}
	if (!temp->upperCell) {
		headCell = headCell->lowerCell;
		Cell<T>* tempDown = temp->lowerCell;
		while (tempDown != nullptr)
		{
			tempDown->upperCell = nullptr;
			tempDown = tempDown->rightCell;
		}
		currCell = active->lowerCell;
	}
	else if (!temp->lowerCell) {
		Cell<T>* tempUp = temp->upperCell;
		while (tempUp != nullptr)
		{
			tempUp->lowerCell = nullptr;
			tempUp = tempUp->upperCell;
		}
		currCell = active->upperCell;
	}
	else {
		Cell<T>* tempUp = temp->upperCell;
		Cell<T>* tempDown = temp->lowerCell;
		while (tempUp != nullptr)
		{
			tempUp->lowerCell = tempDown;
			tempDown->upperCell = tempUp;
			tempUp = tempUp->upperCell;
			tempDown = tempDown->upperCell;
		}
		currCell = active->upperCell;
	}
	numRows--;
}
template <typename T>
void Spreadsheet<T>::clearColumn() {
	Cell<T>* active = currCell;
	while (active->upperCell != nullptr)
	{
		active = active->upperCell;
	}
	while (active) {
		active->value = "0";
		active = active->lowerCell;
	}
}
template <typename T>
void Spreadsheet<T>::clearRow() {
	Cell<T>* active = currCell;
	while (active->leftCell != nullptr)
	{
		active = active->leftCell;
	}
	while (active) {
		active->value = "0";
		active = active->rightCell;
	}
}


template <typename T>
void Spreadsheet<T>::save() {
	ofstream fout;
	fout.open("sheet.txt");
	Cell<T>* temp = headCell;
	Cell<T>* tempRow;
	for (int i = 0; i < numRows; i++) {
		tempRow = temp;
		for (int j = 0; j < numCols; j++) {
			fout << tempRow->value << " ";
			tempRow = tempRow->rightCell;
		}
		fout << endl;
		temp = temp->lowerCell;
	}
	fout.close();
}
template <typename T>
void Spreadsheet<T>::load() {
	ifstream fin;
	fin.open("sheet.txt");
	Cell<T>* temp = headCell;
	Cell<T>* tempRow;
	for (int i = 0; i < numRows; i++) {
		tempRow = temp;
		for (int j = 0; j < numCols; j++) {
			fin >> tempRow->value;
			tempRow = tempRow->rightCell;
		}
		temp = temp->lowerCell;
	}
	fin.close();
}
template <typename T>
void Spreadsheet<T>::resetCellWidth() {
	Cell<T>* temp = headCell;
	Cell<T>* tempRow;
	for (int i = 0; i < numRows; i++) {
		tempRow = temp;
		for (int j = 0; j < numCols; j++) {
			tempRow->value = "0";
			tempRow = tempRow->rightCell;
		}
		temp = temp->lowerCell;
	}
}
template <typename T>
void Spreadsheet<T>::resetCellHeight() {
	Cell<T>* temp = headCell;
	Cell<T>* tempRow;
	for (int i = 0; i < numRows; i++) {
		tempRow = temp;
		for (int j = 0; j < numCols; j++) {
			tempRow->value = "0";
			tempRow = tempRow->lowerCell;
		}
		temp = temp->rightCell;
	}
}
template <typename T>
class ConfidentialSpreadsheetRange {
public:
	Spreadsheet<T>* excel;
	Cell<T>* start;
	Cell<T>* end;
	vector<Cell<T>*>* cells;
	vector<T>* data;
	ConfidentialSpreadsheetRange() {
		excel = nullptr;
		start = nullptr;
		end = nullptr;
		cells = new vector<Cell<T>*>();
	}
	ConfidentialSpreadsheetRange(Spreadsheet<T> sheet) {
		this->excel = sheet;
		start = nullptr;
		end = nullptr;
		cells = new vector<Cell<T>>();
	}
	void populateCellVector();
	int calculateSum();
	void labeler();
	float calculateAverage();
	int calculateCount();
	int calculateMin();
	int calculateMax();
	void copy();
	void cut();
	void paste();

	// Add other member functions as needed
};

bool contains(string data, char c) {
	for (int i = 0; i < data.length(); i++) {
		if (data[i] == c) {
			return true;
		}
	}
	return false;
}
string takeValidData() {
	string data;
	cout << "Enter data: ";
	cin >> data;

	// Check for invalid input (contains comma) or length greater than 4
	while (data.length() > 4 || contains(data, ',')) {
		cout << "Invalid input! Please enter data without a comma and with a length of 4 or less." << endl;
		cout << "Enter data: ";
		cin >> data;
	}

	return data;
}
string validateString(string data) {
	string temp = "";
	if (data.length() > 4) {
		for (int i = 0; i < 4; i++) {
			temp += data[i];
		}
		return temp;
	}
	return data;
}

template <typename T>
void ConfidentialSpreadsheetRange<T>::labeler() {
	int numRows = excel->numRows;
	int numCols = excel->numCols;
	Cell<T>* temp = excel->headCell;
	Cell<T>* currentRow;

	for (int i = 0; i < numRows; i++) {
		currentRow = temp;
		for (int j = 0; j < numCols; j++) {
			currentRow->rowNo = i;
			currentRow->colNo = j;
			currentRow = currentRow->rightCell;
		}
		temp = temp->lowerCell;
	}
}


template <typename T>
void ConfidentialSpreadsheetRange<T>::populateCellVector() {
	if (start->colNo <= end->colNo && start->rowNo >= end->rowNo) {
		Cell<string>* temp = start;
		Cell<string>* currentRow = temp;
		for (int i = start->rowNo; i >= end->rowNo; i--) {
			for (int j = start->colNo; j <= end->colNo; j++) {
				cells->push_back(temp);
				temp = temp->rightCell;
			}
			currentRow = currentRow->upperCell;
			temp = currentRow;
		}
	}
	else if (start->colNo <= end->colNo && start->rowNo < end->rowNo) {
		Cell<string>* temp = start;
		Cell<string>* currentRow = temp;
		for (int i = start->rowNo; i <= end->rowNo; i++) {
			for (int j = start->colNo; j <= end->colNo; j++) {
				cells->push_back(temp);
				temp = temp->rightCell;
			}
			currentRow = currentRow->lowerCell;
			temp = currentRow;
		}
	}
	else if (start->colNo > end->colNo && start->rowNo > end->rowNo) {
		Cell<string>* temp = start;
		Cell<string>* currentRow = temp;
		for (int i = start->rowNo; i >= end->rowNo; i--) {
			for (int j = start->colNo; j >= end->colNo; j--) {
				cells->push_back(temp);
				temp = temp->leftCell;
			}
			currentRow = currentRow->upperCell;
			temp = currentRow;
		}
	}
	else if (start->colNo > end->colNo && start->rowNo < end->rowNo) {
		Cell<string>* temp = start;
		Cell<string>* currentRow = temp;
		for (int i = start->rowNo; i <= end->rowNo; i++) {
			for (int j = start->colNo; j >= end->colNo; j--) {
				cells->push_back(temp);
				temp = temp->leftCell;
			}
			currentRow = currentRow->lowerCell;
			temp = currentRow;
		}
	}
}
template <typename T>
int ConfidentialSpreadsheetRange<T>::calculateSum() {
	string temp = "";
	int sum = 0;
	for (int i = 0; i < cells->size(); i++) {
		temp = to_string(cells->at(i)->value);
		sum += stoi(temp);
	}
	return sum;
}
template <typename T>
float ConfidentialSpreadsheetRange<T>::calculateAverage() {
	return round(((float)calculateSum() / cells->size()) * 10) / 10.0;
}

template <typename T>
int ConfidentialSpreadsheetRange<T>::calculateCount() {
	return cells->size();
}
template <typename T>
int ConfidentialSpreadsheetRange<T>::calculateMin() {
	int min = INT_MAX;
	string temp = "";
	for (int i = 0; i < cells->size(); i++) {
		temp = to_string(cells->at(i)->value);
		if (min > stoi(temp)) {
			min = stoi(temp);
		}
	}
	return min;
}
template <typename T>
int ConfidentialSpreadsheetRange<T>::calculateMax() {
	int max = INT_MIN;
	string temp = "";
	for (int i = 0; i < cells->size(); i++) {
		temp = to_string(cells->at(i)->value);
		if (max < stoi(temp)) {
			max = stoi(temp);
		}
	}
	return max;
}
template <typename T>
void ConfidentialSpreadsheetRange<T>::copy() {
	data = new vector<T>();
	for (int i = 0; i < cells->size(); i++) {
		data->push_back(cells->at(i)->value);
	}
}
template <typename T>
void ConfidentialSpreadsheetRange<T>::cut() {
	data = new vector<T>();
	for (int i = 0; i < cells->size(); i++) {
		data->push_back(cells->at(i)->value);
	}
	for (int i = 0; i < cells->size(); i++) {
		cells->at(i)->value = "0";
	}
}

template <typename T>
void ConfidentialSpreadsheetRange<T>::paste() {
	int index = data->size() - 1;
	if (end->colNo <= start->colNo && end->rowNo <= start->rowNo) {
		Cell<T>* temp = excel->currCell;
		Cell<T>* currentRow = temp;
		for (int i = end->rowNo; i <= start->rowNo; i++) {
			for (int j = end->colNo; j <= start->colNo; j++) {
				temp->value = data->at(index);
				index--;
				if (temp->rightCell == nullptr) {
					excel->makeRightCol();
					excel->numCols++;
				}
				temp = temp->rightCell;
			}
			if (temp->lowerCell == nullptr) {
				excel->makeNewRowAtEnd();
				excel->numRows++;
			}
			currentRow = currentRow->lowerCell;
			temp = currentRow;
		}
	}
	else if (end->colNo <= start->colNo && end->rowNo > start->rowNo) {
		Cell<T>* temp = excel->currCell;
		Cell<T>* currentRow = temp;
		int idx = start->rowNo - end->rowNo + 1;
		int c = 1;
		for (int i = start->rowNo; i <= end->rowNo; i++) {
			for (int j = start->colNo; j <= end->colNo; j++) {
				currentRow->value = data->at(idx - 1);
				idx--;
				if (currentRow->rightCell == nullptr) {
					excel->makeRightCol();
					excel->numCols++;
				}
				currentRow = currentRow->rightCell;
			}
			c++;
			idx = c * (start->colNo - end->colNo + 1);
			if (currentRow->lowerCell == nullptr) {
				excel->makeNewRowAtEnd();
				excel->numRows++;
			}
			temp = temp->lowerCell;
			currentRow = temp;
		}
	}
	else if (end->colNo > start->colNo && end->rowNo > start->rowNo) {
		Cell<T>* temp = excel->currCell;
		Cell<T>* currentRow = temp;
		int idx = 0;
		for (int i = start->rowNo; i <= end->rowNo; i++) {
			for (int j = start->colNo; j <= end->colNo; j++) {
				currentRow->value = data->at(idx);
				idx++;
				if (currentRow->rightCell == nullptr) {
					excel->makeRightCol();
					excel->numCols++;
				}
				currentRow = currentRow->rightCell;
			}
			if (currentRow->lowerCell == nullptr) {
				excel->makeNewRowAtEnd();
				excel->numRows++;
			}
			temp = temp->lowerCell;
			currentRow = temp;
		}
	}
	else if (end->colNo > start->colNo && end->rowNo <= start->rowNo) {
		Cell<T>* temp = excel->currCell;
		Cell<T>* currentRow = temp;
		int c = start->rowNo - end->rowNo;
		int idx = c * (end->colNo - start->colNo + 1);
		for (int i = start->rowNo; i >= end->rowNo; i--) {
			for (int j = start->colNo; j <= end->colNo; j++) {
				currentRow->value = data->at(idx);
				idx++;
				if (currentRow->rightCell == nullptr) {
					excel->makeRightCol();
					excel->numCols++;
				}
				currentRow = currentRow->rightCell;
			}
			c--;
			idx = c * (end->colNo - start->colNo + 1);
			if (currentRow->lowerCell == nullptr) {
				excel->makeNewRowAtEnd();
				excel->numRows++;
			}
			temp = temp->lowerCell;
			currentRow = temp;
		}
	}
}
template <typename T>
class Designing {
public:
	Spreadsheet<T>* excel;

	Designing(Spreadsheet<T> excel) {
		this->excel = excel;
	}
	Designing() {
		excel = nullptr;
	}
	void userManual();
	void ranges(ConfidentialSpreadsheetRange<T>* range);
	void selectRange(Spreadsheet<T>* ss, ConfidentialSpreadsheetRange<T>* range);
	void arrowMovement(string color, bool& modify);
	// Add other member functions as needed
};

template <typename T>
void Designing<T>::userManual() {
	cout << "Press Shift + A to insert row above the selected cell." << endl;
	cout << "Use space to enter data." << endl;
	cout << "Press Shift + U to delete cells by shifting up." << endl;
	cout << "Press Shift + E to delete the selected row." << endl;
	cout << "Press Shift + I to insert cells by shifting right." << endl;
	cout << "Press Shift + D to delete the selected column." << endl;
	cout << "Press Shift + K to insert cells by shifting down." << endl;
	cout << "Press Shift + R to insert a column to the right." << endl;
	cout << "Press Shift + L to insert a column to the left." << endl;
	cout << "Use arrow keys to navigate." << endl;
	cout << "Press Shift + O to delete cells by shifting left." << endl;
	cout << "Press Shift + N to clear the selected row." << endl;
	cout << "Press Shift + M to clear the selected column." << endl;
	cout << "Press Shift + S to save the spreadsheet." << endl;
	cout << "Press Shift + P to load a saved spreadsheet." << endl;
	cout << "Press Tab key to select a range of cells." << endl;
	cout << "Press Shift + W to reset cell width." << endl;
	cout << "Press Shift + X to reset cell height." << endl;
	cout << "Press Shift + Z to set cell width." << endl;
	cout << "Press Shift + H to set cell height." << endl;
	cout << "Use escape to exit." << endl;
	cout << "                           ----------------------------------------------------------------                      " << endl;
}


template <typename T>
void Designing<T>::ranges(ConfidentialSpreadsheetRange<T>* range) {
	string option;
	system("cls");
	cout << "1. Sum" << endl;
	cout << "2. Average" << endl;
	cout << "3. Count" << endl;
	cout << "4. Max" << endl;
	cout << "5. Min" << endl;
	cout << "6. Copy and Paste" << endl;
	cout << "7. Cut and Paste" << endl;
	cout << "0. Back" << endl;
	cout << "Enter option:";
	cin >> option;

	if (option == "1") {
		// sum
		excel->currCell->value = validateString(to_string(range->calculateSum()));
		cout << "Sum of the range is: " << range->calculateSum() << endl;
		Sleep(1000);
	}
	else if (option == "2") {
		// avg
		excel->currCell->value = validateString(to_string(range->calculateAverage()));
		cout << "Average of the range is: " << range->calculateAverage() << endl;
		Sleep(1000);
	}
	else if (option == "3") {
		// count
		excel->currCell->value = validateString(to_string(range->calculateCount()));
		cout << "Count of the range is: " << range->calculateCount() << endl;
		Sleep(1000);
	}
	else if (option == "4") {
		// max
		excel->currCell->value = validateString(to_string(range->calculateMax()));
		cout << "Maximum of the range is: " << range->calculateMax() << endl;
		Sleep(1000);
	}
	else if (option == "5") {
		// min
		excel->currCell->value = validateString(to_string(range->calculateMin()));
		cout << "Minimum of the range is: " << range->calculateMin() << endl;
		Sleep(1000);
	}
	else if (option == "6") {
		// copy
		range->copy();
		range->paste();
		cout << "Range copied" << endl;
		Sleep(1000);
	}
	else if (option == "7") {
		// cut
		range->cut();
		range->paste();
		cout << "Range cut" << endl;
		Sleep(1000);
	}
	else
	{
		cout << "Invalid option" << endl;
		Sleep(1000);
		return;
	}
}
template <typename T>
void Designing<T>::selectRange(Spreadsheet<T>* ss, ConfidentialSpreadsheetRange<T>* range) {
	Cell<T>* current = ss->currCell;
	cout << "Starting cell : " << endl;

	bool selection = true, change = false;
	Sleep(50);
	while (selection) {
		arrowMovement("\33[35m", change);
		if (GetAsyncKeyState(VK_TAB)) {
			if (range->start == nullptr) {
				range->start = ss->currCell;
				cout << "Start cell selected." << endl;
				cout << "Select End cell." << endl;
				change = true;
			}
			else {
				range->end = ss->currCell;
				cout << "End cell selected." << endl;
				selection = false;
				range->start->color = "\33[34m";
				range->end->color = "\33[34m";
			}
		}
		Sleep(100);

		if (change) {
			change = false;
			system("cls");
			userManual();
			ss->displayGrid();
			if (range->start != nullptr) {
				range->start->color = "\33[33m";
				cout << "Start cell selected" << endl;
				cout << "Select end cell" << endl;
			}
			else {
				cout << "Select start cell" << endl;
			}
		}
	}
	ss->currCell = current;
}
template <typename T>
void Designing<T>::arrowMovement(string color, bool& change) {
	if (GetAsyncKeyState(VK_UP) && excel->currCell->upperCell) {
		excel->currCell->color = "\33[37m";
		excel->currCell = excel->currCell->upperCell;
		change = true;
	}
	if (GetAsyncKeyState(VK_RIGHT) && excel->currCell->rightCell) {
		excel->currCell->color = "\33[37m";
		excel->currCell = excel->currCell->rightCell;
		change = true;
	}
	if (GetAsyncKeyState(VK_DOWN) && excel->currCell->lowerCell) {
		excel->currCell->color = "\33[37m";
		excel->currCell = excel->currCell->lowerCell;
		change = true;
	}
	if (GetAsyncKeyState(VK_LEFT) && excel->currCell->leftCell) {
		excel->currCell->color = "\33[37m";
		excel->currCell = excel->currCell->leftCell;
		change = true;
	}
	// Set the color after updating the current cell
	excel->currCell->color = color;
}

void main() {
	Spreadsheet<string>* excel = new Spreadsheet<string>();
	Designing<string>* designing = new Designing<string>();
	designing->excel = excel;
	/*cout << "Do you want to load the previous spreadsheet? (y/n)" << endl;
	string option;
	cin >> option;
	if (option == "y") {
		excel->load();
	}*/
	designing->userManual();
	excel->displayGrid();
	bool running = true, change = false;
	while (running) {
		designing->arrowMovement("\33[35m", change);
		if (GetAsyncKeyState(VK_SPACE)) {
			excel->currCell->value = takeValidData();
			change = true;
		}
		if (GetAsyncKeyState(VK_ESCAPE)) {
			running = false;
			system("cls");
		}
		if (change) {
			system("cls");
			designing->userManual();
			excel->displayGrid();
			change = false;
		}
		else if (GetAsyncKeyState(VK_SHIFT)) {
			if (GetAsyncKeyState('A')) {
				// insert row above
				excel->insertRowAbove();
				change = true;
			}
			else if (GetAsyncKeyState('B')) {
				// insert row below
				excel->insertRowBelow();
				change = true;
			}
			else if (GetAsyncKeyState('R')) {
				// insert column right
				excel->insertColumnToRight();
				change = true;
			}
			else if (GetAsyncKeyState('L')) {
				// insert column left
				excel->insertColumnToLeft();
				change = true;
			}
			else if (GetAsyncKeyState('I')) {
				// insert cell right
				excel->insertCellByRightShift();
				change = true;
			}
			else if (GetAsyncKeyState('K')) {
				// insert cell down
				excel->insertCellByDownShift();
				change = true;
			}
			else if (GetAsyncKeyState('O')) {
				// delete cell left
				excel->deleteCellByLeftShift();
				change = true;
			}
			else if (GetAsyncKeyState('U')) {
				// delete cell up
				excel->deleteCellByUpShift();
				change = true;
			}
			else if (GetAsyncKeyState('D')) {
				// delete column
				excel->deleteColumn();
				change = true;
			}
			else if (GetAsyncKeyState('E')) {
				// delete row
				excel->deleteRow();
				change = true;
			}
			else if (GetAsyncKeyState('M')) {
				// clear column
				excel->clearColumn();
				change = true;
			}
			else if (GetAsyncKeyState('N')) {
				// clear row
				excel->clearRow();
				change = true;
			}
			else if (GetAsyncKeyState('S')) {
				//save
				excel->save();
				change = true;
			}
			else if (GetAsyncKeyState('P')) {
				//load
				excel->load();
				change = true;
			}
			else if (GetAsyncKeyState('W')) {
				// reset cell width
				excel->resetCellWidth();
				change = true;
			}
			else if (GetAsyncKeyState('H')) {
				// reset cell height
				excel->resetCellHeight();
				change = true;
			}
		}
		if (GetAsyncKeyState(VK_TAB)) {
			try {
				ConfidentialSpreadsheetRange<string>* range = new ConfidentialSpreadsheetRange<string>();
				range->excel = excel;
				range->labeler();
				designing->selectRange(excel, range);
				range->populateCellVector();
				while (true) {
					designing->ranges(range);
					if (GetAsyncKeyState(VK_ESCAPE)) {
						break;
					}
					else if (GetAsyncKeyState('0') || GetAsyncKeyState(VK_NUMPAD0)) {
						break;
					}
					Sleep(100);
				}
			}
			catch (exception e) {
				cout << "Error: " << e.what() << endl;
			}
			change = true;
		}
		Sleep(100);
	}
}

