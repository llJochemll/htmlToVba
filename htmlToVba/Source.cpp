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
	string strDBName;
	string strRSTName;

	string strCurrentLine;
	size_t startPos = 0;

	//Get VBA variable names


	while () {
		strHTML.find(from, start_pos)
	}
}

string ReplaceAll(string str, const string& from, const string& to) {
	size_t start_pos = 0;
	while ((start_pos = str.find(from, start_pos)) != std::string::npos) {
		str.replace(start_pos, from.length(), to);
		start_pos += to.length(); // Handles case where 'to' is a substring of 'from'
	}
	return str;
}