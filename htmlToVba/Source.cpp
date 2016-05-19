#include <iostream>
#include <fstream>
#include <sstream>
#include <stdio.h>
#include <algorithm>
#include <string>

using namespace std;

ifstream openFile();
void generateCode(string strHTML);
string ReplaceAll(string str, const string& from, const string& to);

int main() {
	cout << "htmlToVba v0.1" << endl << "by Jochem" << endl;
	system("pause");
	system("cls");

	//Get path
	ifstream fileHTML = openFile();
	//Prepare string
	stringstream bufferHTML;
	string strHTML;
	bufferHTML << fileHTML.rdbuf();
	strHTML = bufferHTML.str();
	//Remove the default newline characters added when file was read
	strHTML.erase(remove_if(strHTML.begin(), strHTML.end(),[](char c) { return c == '\n'; }), strHTML.end());

	cout << strHTML <<endl;

	system("pause");
	system("cls");
	
	generateCode(strHTML);
}

ifstream openFile() {
	while(true) {
		string strPath = "";

		//Ask for path
		cout << "Absolute path to HTML-file: ";
		getline(cin, strPath);

		//Load ifstream object
		ifstream fileHTML(strPath);

		//Check if file exists
		if (fileHTML) {
			cout << "File loaded" << endl;
			return fileHTML;
		}else{
			cout << strPath << " does not exist, try again" << endl;
			system("pause");
			system("cls");
		}
	}
}

void generateCode(string strHTML) {
	//Init vars
	ofstream fileOutput("code.txt");
	string strTBLName;
	string strSQL;
	string strFso;
	string strOutputPath;
	string strCurrentLine;
	string strFldName;
	size_t startPos = 0;
	short int replaceCount = 0;
	long positionStart = 0;
	long positionEnd = 0;

	//Get VBA variable names
	cout << "Name of table to get fields from (Case sensitive): ";
	getline(cin, strTBLName);
	cout << "Path where VBA will put the file (filename and extension included): ";
	getline(cin, strOutputPath);

	//Start writing default code
	cout << endl << "Starting code generation, please wait..." << endl;
	fileOutput << "Dim strSQL As String" << endl;
	fileOutput << "Dim myDB As DAO.Database" << endl;
	fileOutput << "Dim myRST As DAO.Recordset" << endl;
	fileOutput << "Dim fso As Object" << endl;
	fileOutput << "Dim Fileout As Object" << endl; 
	fileOutput << "Set myDB = CurrentDb" << endl;
	strSQL = "strSQL = \"SELECT * FROM ";
	strSQL.append(strTBLName);
	strSQL.append("\"");
	fileOutput << strSQL << endl;
	fileOutput << "Set myRST = myDB.OpenRecordset(strSQL, dbOpenDynaset, dbSeeChanges)" << endl;
	fileOutput << "If myRST.BOF And myRST.EOF Then" << endl;
	fileOutput << "Exit Sub" << endl;
	fileOutput << "End If" << endl;
	fileOutput << "Set fso = CreateObject(\"Scripting.FileSystemObject\")" << endl;
	strFso = "Set Fileout = fso.CreateTextFile(\"";
	strFso.append(strOutputPath);
	strFso.append("\", True, True)");
	fileOutput << strFso << endl;
	fileOutput << "myRST.MoveFirst" << endl;
	fileOutput << "Do While Not myRST.EOF" << endl;

	//Start writing dynamic code
	while (strHTML.find("@@@", startPos) != string::npos) {
		strCurrentLine = "Fileout.Write \"";
		positionStart = strHTML.find("@@@", startPos) + 3;
		positionEnd = strHTML.find("@@@", positionStart);
		strFldName = strHTML.substr(positionStart, positionEnd - positionStart);
		strCurrentLine.append(strHTML.substr(0, positionStart - 3));
		strCurrentLine.append("\"");
		strCurrentLine.append(" + str(myRST(\"");
		strCurrentLine.append(strFldName);
		strCurrentLine.append("\"))");
		fileOutput << strCurrentLine << endl;
		strHTML = strHTML.substr(positionEnd + 3, strHTML.length() + 1);

		replaceCount++;
	}
	strCurrentLine = "Fileout.Write \"";
	strCurrentLine.append(strHTML);
	strCurrentLine.append("\"");
	fileOutput << strCurrentLine << endl;

	//Fileout.Write "your string goes here"

	//Continue writing default code
	fileOutput << "myRST.MoveNext" << endl;
	fileOutput << "Loop" << endl;

	//Save file
	fileOutput.close();

	cout << "Replaced " << replaceCount << " strings with fields" << endl;

	system("pause");
}

string ReplaceAll(string str, const string& from, const string& to) {
	size_t start_pos = 0;
	while ((start_pos = str.find(from, start_pos)) != std::string::npos) {
		str.replace(start_pos, from.length(), to);
		start_pos += to.length(); // Handles case where 'to' is a substring of 'from'
	}
	return str;
}