PACKAGE CREOXL IS
/*********************************************************************************
//   	 _____ _______ ______ ______ __ ___ __		     ___      ______
// 	|     |  __   |	  ___|	    |   |  |  |       __ __ |_  |    |___   |
// 	|   --|	     -|   ___|	|   |-    -|  |__    |  |  | _| |_ __	 |  |
//	|_____|___|___|______|______|__|___|_____|    \___/ |_____|__|	 |__|
//
//********************************************************************************
//  CREOXL v1.7 
//  
//********************************************************************************
//
//  BRIEF:
//	This package can programmatically produce a formatted excel file and also
//	import contents of excel file into a block or table.
//  
//  REQUIREMENTS:
//	WPSOffice or Excel and set either as default program.
//
//  HOW TO IMPORT:
// 	Copy the following to your project:
//	1. CREOXL (Package Spec)
//	2. CREOXL (Package Body)
//	3. IMPORT_CUSTOM_BLOCK (Implementation is found in comments)
//	4. IMPORT_CUSTOM_TABLE (Implementation is found in comments)
//
//  HOW TO USE:
//
//  WHEN GENERATING EXCEL FILE:
//  	1. Start with CREOXL.INITIALIZE;
//	2. Format your excel file using the available functions within this library
//	   to fit your needs. 
//  	3. Use CREOXL.DISPLAY_OUTPUT; or SAVEAS('YourPath')
//	   NOTE: The former will open an untitled formatted excel file
//	   and the latter saves the formatted file to the specified path 
//	4. End with CREOXL.TERMINATE; 
//	   NOTE: This will release all the handles and prevent memory leaks
//
//  WHEN IMPORTING:
//  	1. You can use IMPORT_TO_BLOCK or IMPORT_TO_TABLE
//	   NOTE: 
// 	   When using IMPORT_TO_BLOCK use IMPORT_CUSTOM_BLOCK function
//	   to assign values to your block. DO NOT edit the functions within
//	   this library unless you know what you are doing.
//	   When using IMPORT_TO_TABLE use IMPORT_CUSTOM_TABLE function
//	   to map fields and values.
//	2. You can write your own custom importing function by using the following
//	   guide:
//		IMPORT_INITIALIZE
//		--
//		write your loop here and use GET_TEXT to fetch the data
//		--
//		IMPORT_TERMINATE
//
//  CHANGE LOGS:
//  	v1.0 Initial release
//  	v1.1 Now supports importing
//	     changed RELEASE_OBJECTS to TERMINATE
//	     added INITIALIZE_IMPORT
//	     added INITIALIZE_TERMINATE
//	     added GET_TEXT
//	v1.2 added SAVEAS function
//	v1.3 added SET_ORIENTATION function
//	v1.4 added SET_IMAGE function
//	v1.5 added IMPORT_TO_BLOCK function
//	v1.6 added IMPORT_TO_TABLE function
//	     changed import implementation
//	v1.7 exposed BORDER_WDITH
//
**********************************************************************************/

	-- HANDLES
	APPLICATION OLE2.OBJ_TYPE;
	WORKBOOKS   OLE2.OBJ_TYPE;
	WORKBOOK    OLE2.OBJ_TYPE;
	WORKSHEETS  OLE2.OBJ_TYPE;
	WORKSHEET   OLE2.OBJ_TYPE;
	PAGESETUP		OLE2.OBJ_TYPE;
	CELL        OLE2.OBJ_TYPE;
	FONT        OLE2.OBJ_TYPE;
	BORDER 			OLE2.OBJ_TYPE;
	RANGE       OLE2.OBJ_TYPE;
	RANGE_COL   OLE2.OBJ_TYPE;
	COLUMN_NUM	OLE2.OBJ_TYPE;
	ROW_NUM			OLE2.OBJ_TYPE;
 	ARGS 				OLE2.LIST_TYPE;
 	PICTURE 	  OLE2.OBJ_TYPE;
	PICTURES 		OLE2.OBJ_TYPE;
	BORDER_WIDTH NUMBER DEFAULT 2;

	-- PROCEDURES FOR EXPORTING
	PROCEDURE INITIALIZE;
	PROCEDURE TERMINATE;
	PROCEDURE DISPLAY_OUTPUT;
	PROCEDURE AUTOFIT;
	PROCEDURE MOVE_ACTIVESHEET (P_PAGE NUMBER);
	PROCEDURE MERGE_CELLS(P_ROW IN VARCHAR2, P_COLUMN IN VARCHAR2);
	PROCEDURE SET_TEXT(P_ROW IN VARCHAR2, P_COLUMN IN VARCHAR2, P_VALUE IN VARCHAR2);
	PROCEDURE SET_FONTSTYLE(P_ROW IN VARCHAR2, P_COLUMN IN VARCHAR2, P_STYLE IN VARCHAR2);
	PROCEDURE SET_FONTSIZE(P_ROW IN VARCHAR2, P_COLUMN IN VARCHAR2, P_SIZE IN NUMBER);
	PROCEDURE SET_FONTNAME(P_ROW IN VARCHAR2, P_COLUMN IN VARCHAR2, P_FONTNAME IN VARCHAR2);
	PROCEDURE SET_FONTCOLOR(P_ROW IN VARCHAR2, P_COLUMN IN VARCHAR2, P_FONTCOLOR IN NUMBER);
	PROCEDURE SET_BACKCOLOR(P_ROW IN VARCHAR2, P_COLUMN IN VARCHAR2, P_FONTCOLOR IN NUMBER);
	PROCEDURE SHEET_NAME (P_SHEET_NAME IN VARCHAR2);
	PROCEDURE SET_BORDER(P_ROW IN VARCHAR2, P_COLUMN IN VARCHAR2, P_AXIS IN VARCHAR2);
	PROCEDURE SET_COLUMN_WIDTH(P_COLUMN IN NUMBER, P_WIDTH IN NUMBER);                 
	PROCEDURE SET_ROW_HEIGHT(P_ROW IN NUMBER, P_HEIGHT IN NUMBER);
	PROCEDURE SET_CENTER(P_ROW IN VARCHAR2, P_COLUMN IN VARCHAR2);                     
	PROCEDURE SET_TEXTWRAP(P_ROW IN VARCHAR2, P_COLUMN IN VARCHAR2);
	PROCEDURE SET_FULLBORDER(P_ROW IN VARCHAR2, P_COLUMN IN VARCHAR2);
	PROCEDURE SET_FULLBORDER_RANGE(P_ROW IN NUMBER, P_START_COL IN NUMBER, 
						P_END_COL IN NUMBER);
	PROCEDURE SET_BORDER_RANGE(P_ROW IN NUMBER, P_START_COL IN NUMBER, 
						P_END_COL IN NUMBER, P_AXIS IN VARCHAR2);
	PROCEDURE SAVEAS(P_PATHNAME IN VARCHAR2);
	PROCEDURE SET_ORIENTATION(P_ORIENTATION IN NUMBER);
	PROCEDURE SET_IMAGE(P_ROW IN NUMBER, P_COLUMN IN NUMBER, P_FILEPATH IN VARCHAR2, 
						P_HEIGHT IN NUMBER, P_WIDTH IN NUMBER);

	-- PROCEDURES FOR IMPORTING						
	PROCEDURE IMPORT_INITIALIZE(P_PATHNAME IN VARCHAR2);
	PROCEDURE IMPORT_TERMINATE;
	FUNCTION IMPORT_TO_BLOCK(P_PATHNAME IN VARCHAR2, P_BLOCK_NAME IN VARCHAR2, 
		P_ROW_START IN NUMBER, P_COLUMN_COUNT IN NUMBER, P_SUCCESS IN VARCHAR2, 
		P_FAIL IN VARCHAR2, ROW_COUNTER OUT NUMBER, P_CUSTOM_MODE BOOLEAN)RETURN VARCHAR2;
	FUNCTION IMPORT_TO_TABLE(P_PATHNAME IN VARCHAR2, P_TABLE_NAME IN VARCHAR2, P_ROW_START IN NUMBER, 
		P_COLUMN_COUNT IN NUMBER, P_BUFFER_SIZE IN NUMBER) RETURN NUMBER;
	FUNCTION  GET_TEXT(P_ROW IN NUMBER, P_COLUMN IN NUMBER) RETURN VARCHAR2;
END;
