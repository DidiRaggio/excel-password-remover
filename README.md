# Excel Password Remover

## Os: Windows

## Python version: 3.6.1

## Dependencies:
	os 
	sys
	glob
	win32com.client

## Usage:
	1) add excel files to excel_files folder in working directory.
	2) a) with default password = 'password', in terminal run:
						$ python remove_excel_password.py
	   b) with inputed password, in terminal run:
						$ python remove_excel_password.py <INPUTTED PASSWORD>