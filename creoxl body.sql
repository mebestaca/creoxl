PACKAGE BODY CREOXL IS
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

	/**********************************************************************************
	// EXPORT PROCEDURES
	**********************************************************************************/

	/**********************************************************************************
	// Takes nothing returns nothing
	// -	Displays the opens the formatted excel file then releases all handles
	**********************************************************************************/
  	PROCEDURE INITIALIZE IS
	BEGIN
			APPLICATION:=OLE2.CREATE_OBJ('EXCEL.APPLICATION');
			WORKBOOKS:= OLE2.GET_OBJ_PROPERTY(APPLICATION, 'WORKBOOKS');
			WORKBOOK := OLE2.INVOKE_OBJ(WORKBOOKS,'ADD');
			WORKSHEETS:= OLE2.GET_OBJ_PROPERTY(WORKBOOK, 'WORKSHEETS');
			WORKSHEET := OLE2.GET_OBJ_PROPERTY(application,'activesheet');
	END;
	
	/**********************************************************************************
	// Takes nothing returns nothing
	// -	Clears all the handles and display the formatted excel window
	**********************************************************************************/
	PROCEDURE TERMINATE IS
	BEGIN
			OLE2.RELEASE_OBJ(WORKSHEETS);
			OLE2.RELEASE_OBJ(WORKBOOKS);
			OLE2.RELEASE_OBJ(WORKSHEET);
			OLE2.RELEASE_OBJ(WORKBOOK);
			OLE2.RELEASE_OBJ(APPLICATION);
	END;
	
	/**********************************************************************************
	// Takes nothing returns nothing
	// -	Opens an unsaved untitled excel file
	**********************************************************************************/
	PROCEDURE DISPLAY_OUTPUT IS
	BEGIN
			OLE2.SET_PROPERTY(APPLICATION, 'VISIBLE', 'TRUE');
	END;

	/**********************************************************************************
	// Takes nothing returns nothing
	// -	Allegedly prevents text overflow
	// Note:
	// -	Details are unclear. Ignore it.
	**********************************************************************************/
	PROCEDURE AUTOFIT IS
	BEGIN
			CREOXL.RANGE := OLE2.GET_OBJ_PROPERTY(WORKSHEET,'USEDRANGE');
			CREOXL.RANGE_COL := OLE2.GET_OBJ_PROPERTY(RANGE,'COLUMNS');
			OLE2.INVOKE(RANGE_COL,'AUTOFIT');
			OLE2.RELEASE_OBJ(RANGE );
			OLE2.RELEASE_OBJ(RANGE_COL );
	END;
	
	/**********************************************************************************
	// Takes page number returns nothing
	// -	Sets the page active
	**********************************************************************************/
	PROCEDURE MOVE_ACTIVESHEET (P_PAGE number) IS
	BEGIN
			ARGS := OLE2.CREATE_ARGLIST;
			OLE2.ADD_ARG(ARGS, p_page);
			WORKSHEET := OLE2.GET_OBJ_PROPERTY(WORKBOOK, 'Worksheets', args);
			OLE2.DESTROY_ARGLIST(args);
			OLE2.INVOKE(WORKSHEET, 'Select');
	END;
	
	/**********************************************************************************
	// Takes row and column values returns nothing
	// -	Merges cells within the range of row and column values specified
	// Usage
	// -	CREOXL.MERGE_CELLS(1, 'A:B'); Result: A1 and B1 merged
	// -	CREOXL.MERGE_CELLS('1:2', 'A'); Result: A1 and A2 merged
	// - 	CREOXL.MERGE_CELLS('1:2', 'A:B'); Result: A1, A2, B1, B2 merged
	**********************************************************************************/
	PROCEDURE MERGE_CELLS(P_ROW IN VARCHAR2,P_COLUMN IN VARCHAR2) IS
	BEGIN
			ARGS := OLE2.CREATE_ARGLIST;
			OLE2.ADD_ARG(ARGS, P_COLUMN);
			COLUMN_NUM := OLE2.GET_OBJ_PROPERTY(WORKSHEET, 'Columns', args);
			OLE2.DESTROY_ARGLIST(ARGS);
			ARGS := OLE2.CREATE_ARGLIST;
			OLE2.ADD_ARG(args, P_ROW);
			ROW_NUM := OLE2.GET_OBJ_PROPERTY(COLUMN_NUM, 'Rows', args);
			OLE2.DESTROY_ARGLIST(args);
			OLE2.INVOKE(ROW_NUM, 'Merge');
	END;
	
	/**********************************************************************************
	// Takes row, column values and the text value returns nothing
	// -	Puts P_VALUE on target cell represented by row and column values
	// Usage
	// - 	CREOXL.SET_TEXT(1,1,'A'); Result: A1 now contains A
	**********************************************************************************/
	PROCEDURE SET_TEXT(P_ROW IN VARCHAR2, P_COLUMN IN VARCHAR2, P_VALUE IN VARCHAR2) IS
	BEGIN
			ARGS:=OLE2.CREATE_ARGLIST;
			OLE2.ADD_ARG(ARGS, P_ROW   );
			OLE2.ADD_ARG(ARGS, P_COLUMN);
			CELL:=OLE2.GET_OBJ_PROPERTY(WORKSHEET, 'CELLS', ARGS);
			OLE2.DESTROY_ARGLIST(ARGS);
			OLE2.SET_PROPERTY(CELL, 'VALUE', P_VALUE);
			OLE2.RELEASE_OBJ(CELL);  
	END;
	
	/**********************************************************************************
	// Takes row, column values and text style returns nothing
	// -	Applies P_STYLE on target cell represented by row and column values
	// Usage
	// -	CREOXL.SET_FONTSTYLE(1,1,'BOLD'); Result: A1 is now bold
	// -  	CREOXL.SET_FONTSTYLE(1,1,'ITALIC'); Result: A1 is now italic
	// -	CREOXL.SET_FONTSTYLE(1,1,'UNDERLINE'); Result: A1 is now underlined
	**********************************************************************************/
	PROCEDURE SET_FONTSTYLE(P_ROW IN VARCHAR2, P_COLUMN IN VARCHAR2, P_STYLE IN VARCHAR2) IS
	BEGIN
			ARGS:=OLE2.CREATE_ARGLIST;
			OLE2.ADD_ARG(ARGS, P_ROW   );
			OLE2.ADD_ARG(ARGS, P_COLUMN);
			CELL:=OLE2.GET_OBJ_PROPERTY(WORKSHEET, 'CELLS', ARGS);
			FONT := OLE2.GET_OBJ_PROPERTY (CELL, 'FONT');
			OLE2.SET_PROPERTY (FONT, UPPER(P_STYLE), TRUE);
			OLE2.SET_PROPERTY (FONT, UPPER(P_STYLE), 2);
			OLE2.RELEASE_OBJ(FONT);
			OLE2.RELEASE_OBJ(CELL);
			OLE2.DESTROY_ARGLIST(ARGS);
	END;
	
	/**********************************************************************************
	// Takes row, column values and size returns nothing
	// -	Alters the cell font size into P_SIZE of the target cell represented by row and 
	//		column values
	// Usage
	// - 	CREOXL.SET_FONTSIZE(1,1,80); Result: A1 now has a font size of 80
	**********************************************************************************/
	PROCEDURE SET_FONTSIZE(P_ROW IN VARCHAR2, P_COLUMN IN VARCHAR2, P_SIZE IN NUMBER) IS
	BEGIN
			ARGS:=OLE2.CREATE_ARGLIST;
			OLE2.ADD_ARG(ARGS, P_ROW   );
			OLE2.ADD_ARG(ARGS, P_COLUMN);
			CELL:=OLE2.GET_OBJ_PROPERTY(WORKSHEET, 'CELLS', ARGS);
			FONT := OLE2.GET_OBJ_PROPERTY (CELL, 'FONT');
			OLE2.SET_PROPERTY (FONT, 'SIZE', P_SIZE);
			OLE2.RELEASE_OBJ(FONT);
			OLE2.RELEASE_OBJ(CELL);
			OLE2.DESTROY_ARGLIST(ARGS);
	END;
	
	/**********************************************************************************
	// Takes row, column values and font name returns nothing
	// -	Changes the font of the target cell represented by row and column values
	// Usage:
	// -	CREOXL.SET_FONTNAME(1,1,'Calibri'); Result: A1 now uses Calibri
	**********************************************************************************/
	PROCEDURE SET_FONTNAME(P_ROW IN VARCHAR2, P_COLUMN IN VARCHAR2, P_FONTNAME IN VARCHAR2) IS
	BEGIN
			ARGS:=OLE2.CREATE_ARGLIST;
			OLE2.ADD_ARG(ARGS, P_ROW   );
			OLE2.ADD_ARG(ARGS, P_COLUMN);
			CELL:=OLE2.GET_OBJ_PROPERTY(WORKSHEET, 'CELLS', ARGS);
			FONT := OLE2.GET_OBJ_PROPERTY (CELL, 'FONT');
			OLE2.SET_PROPERTY (FONT, 'NAME', P_FONTNAME);
			OLE2.RELEASE_OBJ(FONT);
			OLE2.RELEASE_OBJ(CELL);
			OLE2.DESTROY_ARGLIST(ARGS);
	END;
	
	/**********************************************************************************
	// Takes row, column values and font color returns nothing
	// -	Changes the font color of the target cell represented by row and column values
	// Usage:
	// -	CREOXL.SET_FONTNAME(1,1,1); Result: A1 now uses a different color
	// Note:
	// -	Color index is not experimented enough. Feel free to add color values under
	//	'Known Colors' section
	// Known Colors:
	// - 	None
	**********************************************************************************/
	PROCEDURE SET_FONTCOLOR(P_ROW IN VARCHAR2, P_COLUMN IN VARCHAR2, P_FONTCOLOR IN NUMBER) IS
	BEGIN
			ARGS:=OLE2.CREATE_ARGLIST;
			OLE2.ADD_ARG(ARGS, P_ROW   );
			OLE2.ADD_ARG(ARGS, P_COLUMN);
			CELL:=OLE2.GET_OBJ_PROPERTY(WORKSHEET, 'CELLS', ARGS);
			FONT := OLE2.GET_OBJ_PROPERTY (CELL, 'FONT');
			OLE2.SET_PROPERTY (FONT, 'COLORINDEX', P_FONTCOLOR);
			OLE2.RELEASE_OBJ(FONT);
			OLE2.RELEASE_OBJ(CELL);
			OLE2.DESTROY_ARGLIST(ARGS);
	END;
	
	/**********************************************************************************
	// Takes row, column values and backcolor returns nothing
	// -	Changes the font of the target cell represented by row and column values
	// Usage:
	// -	CREOXL.SET_FONTNAME(1,1,1); Result: A1 now uses a different color
	// Note:
	// -	Color index is not experimented enough. Feel free to add color values under
	//	'Known Colors' section
	// Known Colors:
	// - 	None
	**********************************************************************************/
	PROCEDURE SET_BACKCOLOR(P_ROW IN VARCHAR2, P_COLUMN IN VARCHAR2, P_FONTCOLOR IN NUMBER) IS
	BEGIN
			ARGS:=OLE2.CREATE_ARGLIST;
			OLE2.ADD_ARG(ARGS, P_ROW   );
			OLE2.ADD_ARG(ARGS, P_COLUMN);
			CELL := OLE2.GET_OBJ_PROPERTY(WORKSHEET, 'CELLS', ARGS);
			FONT := OLE2.GET_OBJ_PROPERTY (CELL, 'INTERIOR');
			OLE2.SET_PROPERTY (FONT, 'COLORINDEX', P_FONTCOLOR);
			OLE2.RELEASE_OBJ(FONT);
			OLE2.RELEASE_OBJ(CELL);
			OLE2.DESTROY_ARGLIST(ARGS);
	END;	
	
	/**********************************************************************************
	// Takes sheet name returns nothing
	// -	Alters the name of the current sheet
	**********************************************************************************/
	PROCEDURE SHEET_NAME (P_SHEET_NAME IN VARCHAR2) IS
	BEGIN
 			OLE2.SET_PROPERTY(WORKSHEET, 'Name', P_SHEET_NAME);
	END;
	
	/**********************************************************************************
	// Takes row, column and axis returns nothing
	// -	Applies a border on target cell represented by the row and column values to
	//		the defined axis
	// Usage:
	// -	CREOXL.SET_BORDER(1,1,'LEFT'); Result: A1 now has a left border
	// -	CREOXL.SET_BORDER(1,1,'RIGHT'); Result: A1 now has a right border
	// -	CREOXL.SET_BORDER(1,1,'TOP'); Result: A1 now has a top border
	// -	CREOXL.SET_BORDER(1,1,'BOTTOM'); Result: A1 now has a bottom border
	// Note:
	// -	Line style and line color are available but arent exposed
	// -	To adjust the border width use
	//		CREOXL.BORDER_WIDTH := 4; the default line width is 2
	// -	To expose the hidden fields just add a global variable to the package
	//		spec, add a default value and adjust it before calling the function
	//		use this approach to avoid adjusting the original functions and procedures
	//		so that its still backwards compatible to the older versions
	**********************************************************************************/
	PROCEDURE SET_BORDER(P_ROW IN VARCHAR2, P_COLUMN IN VARCHAR2, P_AXIS IN VARCHAR2) IS
			axis number := 0;
	BEGIN
			ARGS:=OLE2.CREATE_ARGLIST;
			OLE2.ADD_ARG(ARGS, P_ROW   );
			OLE2.ADD_ARG(ARGS, P_COLUMN);
			CELL:=OLE2.GET_OBJ_PROPERTY(WORKSHEET, 'CELLS', ARGS);
			OLE2.DESTROY_ARGLIST(ARGS);	
			IF (P_AXIS = 'LEFT') THEN
				axis := 7;
			ELSIF (P_AXIS = 'TOP') THEN
				axis := 8;
			ELSIF (P_AXIS = 'BOTTOM') THEN
				axis := 9;
			ELSIF (P_AXIS = 'RIGHT') THEN
				axis := 10;
			END IF;
			ARGS:=OLE2.CREATE_ARGLIST;
			OLE2.ADD_ARG(ARGS, axis);
			BORDER := OLE2.GET_OBJ_PROPERTY(CELL, 'Borders', ARGS);
			OLE2.SET_PROPERTY (BORDER, 'LineStyle', 1); -- CONTINOUS
			OLE2.SET_PROPERTY (BORDER, 'Weight', BORDER_WIDTH); -- 2 THIN -- 4 THICK
			OLE2.SET_PROPERTY (BORDER, 'ColorIndex', -4105);
			OLE2.DESTROY_ARGLIST(ARGS);
			OLE2.RELEASE_OBJ(CELL);
			OLE2.RELEASE_OBJ(BORDER);
	END;
	
	/**********************************************************************************
	// Takes row, column and width values returns nothing
	// -	Alters the width of the target cell represented by row and column values
	// Usage:
	// - 	CREOXL.SET_COLUMN_WIDTH(1,1); Result: A now as a width of 1px
	**********************************************************************************/
	PROCEDURE SET_COLUMN_WIDTH(P_COLUMN IN NUMBER, P_WIDTH IN NUMBER) IS
	BEGIN
			ARGS := OLE2.CREATE_ARGLIST;
			OLE2.ADD_ARG(ARGS, P_COLUMN); 
			COLUMN_NUM := OLE2.GET_OBJ_PROPERTY(WORKSHEET, 'Columns', args);
			OLE2.DESTROY_ARGLIST(ARGS);
			OLE2.SET_PROPERTY(COLUMN_NUM , 'ColumnWidth', P_WIDTH);
			OLE2.RELEASE_OBJ(COLUMN_NUM);
	END;
	
	/**********************************************************************************
	// Takes row, column and height values returns nothing
	// -	Alters the height of the target cell represented by row and column values
	// Usage:
	// - 	CREOXL.SET_ROW_HEIGHT(1,1); Result: row 1 now as a height of 1px
	**********************************************************************************/
	PROCEDURE SET_ROW_HEIGHT(P_ROW IN NUMBER, P_HEIGHT IN NUMBER) IS
	BEGIN
			ARGS := OLE2.CREATE_ARGLIST;
			OLE2.ADD_ARG(ARGS, P_ROW); 
			COLUMN_NUM := OLE2.GET_OBJ_PROPERTY(WORKSHEET, 'Rows', args);
			OLE2.DESTROY_ARGLIST(ARGS);
			OLE2.SET_PROPERTY(COLUMN_NUM , 'RowHeight', P_HEIGHT);
			OLE2.RELEASE_OBJ(COLUMN_NUM);
	END;

	/**********************************************************************************
	// Takes row and column returns nothing
	// -	Centers the text of the target cell represented by row and column
	**********************************************************************************/
	PROCEDURE SET_CENTER(P_ROW IN VARCHAR2, P_COLUMN IN VARCHAR2) IS
	BEGIN
			ARGS := OLE2.CREATE_ARGLIST;
			OLE2.ADD_ARG(ARGS, P_ROW   );
			OLE2.ADD_ARG(ARGS, P_COLUMN);
			COLUMN_NUM := OLE2.GET_OBJ_PROPERTY(WORKSHEET, 'Cells', args);
			OLE2.DESTROY_ARGLIST(ARGS);
			OLE2.SET_PROPERTY(COLUMN_NUM , 'HorizontalAlignment', -4108);-- -4108);
			OLE2.RELEASE_OBJ(COLUMN_NUM);
	END;	
	
	/**********************************************************************************
	// Takes row and column returns nothing
	// -	Applies textwrap on the target cell represented by row and column
	**********************************************************************************/
	PROCEDURE SET_TEXTWRAP(P_ROW IN VARCHAR2, P_COLUMN IN VARCHAR2) IS
	BEGIN
			ARGS := OLE2.CREATE_ARGLIST;
			OLE2.ADD_ARG(ARGS, P_ROW   );
			OLE2.ADD_ARG(ARGS, P_COLUMN);
			COLUMN_NUM := OLE2.GET_OBJ_PROPERTY(WORKSHEET, 'Cells', args);
			OLE2.DESTROY_ARGLIST(ARGS);
			OLE2.SET_PROPERTY(COLUMN_NUM , 'Wraptext', 'true');-- -4108);
			OLE2.RELEASE_OBJ(COLUMN_NUM);
	END;
	
	/**********************************************************************************
	// Some nice wrapper functions
	**********************************************************************************/
	
	/**********************************************************************************
	// Takes row and column returns nothing
	// -	Applies full border on the target cell represented by row and column
	**********************************************************************************/
	PROCEDURE SET_FULLBORDER(P_ROW IN VARCHAR2, P_COLUMN IN VARCHAR2) IS
	BEGIN
			SET_BORDER(P_ROW, P_COLUMN, 'TOP');
			SET_BORDER(P_ROW, P_COLUMN, 'BOTTOM');
			SET_BORDER(P_ROW, P_COLUMN, 'LEFT');
			SET_BORDER(P_ROW, P_COLUMN, 'RIGHT');
	END;
	
	/**********************************************************************************
	// Takes row, column and end_column returns nothing
	// -	Applies full border on the target cell represented by row and column
	// 	also applies full border on the next cells represented by end_column 
	**********************************************************************************/
	PROCEDURE SET_FULLBORDER_RANGE(P_ROW IN NUMBER, P_START_COL IN NUMBER, P_END_COL IN NUMBER) IS
			i number;
	BEGIN
			FOR i IN P_START_COL..P_END_COL LOOP
				CREOXL.SET_FULLBORDER(P_ROW, i);
			END LOOP;
	END;
	
	/**********************************************************************************
	// Takes row, column and end_column returns nothing
	// -	Applies border on the target cell represented by row and column to the defined axis
	// 	also applies full border on the next cells represented by end_column 
	**********************************************************************************/
	PROCEDURE SET_BORDER_RANGE(P_ROW IN NUMBER, P_START_COL IN NUMBER, P_END_COL IN NUMBER, P_AXIS IN VARCHAR2) IS
			i number;
	BEGIN
			FOR i IN P_START_COL..P_END_COL LOOP
				CREOXL.SET_BORDER(P_ROW, i, P_AXIS);
			END LOOP;
	END;
	
	/**********************************************************************************
	// Takes pathname returns nothing
	// -	saves the file on the specified path
	// -	include the filename at the end of the path when saving
	//		example:
	//		CREOXL.SAVEAS('c:\csv\wassup.xls');
	// -	dont use DISPLAY_OUTPUT when using SAVEAS
	**********************************************************************************/
	PROCEDURE SAVEAS(P_PATHNAME IN VARCHAR2) IS
	BEGIN
			ARGS := OLE2.CREATE_ARGLIST;
			OLE2.ADD_ARG(ARGS, P_PATHNAME);
			OLE2.INVOKE(WORKBOOK, 'SaveAs', ARGS);
			OLE2.DESTROY_ARGLIST(ARGS);
	END;
	
	
	/**********************************************************************************
	// Takes orientation returns nothing
	// -	1 for portrait
	// -	2 for landscape
	**********************************************************************************/
	PROCEDURE SET_ORIENTATION(P_ORIENTATION IN NUMBER) IS
	BEGIN
			ARGS := OLE2.CREATE_ARGLIST;
			PAGESETUP := OLE2.GET_OBJ_PROPERTY(WORKSHEET, 'PageSetup', args);
			OLE2.DESTROY_ARGLIST(ARGS);
			OLE2.SET_PROPERTY(PAGESETUP , 'Orientation', P_ORIENTATION);-- -4108);
			OLE2.RELEASE_OBJ(PAGESETUP);
	END;
	
	/**********************************************************************************
	// Takes row, column, path, height and width
	// -	Does not get inserted into a cell but instead the top and left
	//	  of the image starts at the top left corner of the specified
	//	  row and column
	**********************************************************************************/
	PROCEDURE SET_IMAGE(P_ROW IN NUMBER, P_COLUMN IN NUMBER, P_FILEPATH IN VARCHAR2, P_HEIGHT IN NUMBER, P_WIDTH IN NUMBER) IS
	BEGIN
			ARGS := OLE2.CREATE_ARGLIST;
			OLE2.ADD_ARG(ARGS, P_ROW);
			OLE2.ADD_ARG(ARGS, P_COLUMN);
			CELL := OLE2.GET_OBJ_PROPERTY(WORKSHEET, 'Cells', args);
			OLE2.DESTROY_ARGLIST(ARGS);
			PICTURES := OLE2.INVOKE_obj(WORKSHEET,'Pictures');
			ARGS := OLE2.CREATE_ARGLIST;
			OLE2.ADD_ARG(args, P_FILEPATH);
			PICTURE := OLE2.INVOKE_obj(PICTURES,'Insert',args);
			OLE2.SET_PROPERTY(picture,'Height' , P_HEIGHT);
			OLE2.SET_PROPERTY(picture,'Width' , P_WIDTH);
			OLE2.SET_PROPERTY(picture,'Left' , ole2.get_num_property (cell, 'Left'));
			OLE2.SET_PROPERTY(picture,'Top' , ole2.get_num_property (cell, 'Top'));
			OLE2.DESTROY_ARGLIST(ARGS);
			OLE2.RELEASE_OBJ(PICTURE);
			OLE2.RELEASE_OBJ(CELL);
	END;
	
	/**********************************************************************************
	// IMPORT PROCEDURES
	**********************************************************************************/
	
	/**********************************************************************************
	// Takes pathname returns nothing
	// -	Use GET_FILE_NAME(directory_name => 'C:\',message => 'Select filename to Open.')
	// 	command to safely find valid pathnames. This procedure will open a wps process
	// 	so its imperative that you run IMPORT_TERMINATE after reading the file or it
	// 	will create a new process everytime you read a file. 
	//
	// Note
	// -	Dont forget to run IMPORT_TERMINATE after reading a file to avoid memory
	// 	leaks
	**********************************************************************************/
	PROCEDURE IMPORT_INITIALIZE(P_PATHNAME IN VARCHAR2) IS
	BEGIN
			APPLICATION := OLE2.create_obj('Excel.Application');  
			OLE2.set_property(application,'Visible','false');    
			OLE2.set_property(application,'DisplayAlerts','false');
			WORKBOOKS := OLE2.Get_Obj_Property(application, 'Workbooks');
			ARGS := OLE2.CREATE_ARGLIST;
			OLE2.add_arg(args,P_PATHNAME);
			WORKBOOK := OLE2.GET_OBJ_PROPERTY(workbooks,'Open',args); 
			OLE2.destroy_arglist(args);
			WORKSHEETS := OLE2.GET_OBJ_PROPERTY(workbook, 'Worksheets');
			WORKSHEET := OLE2.GET_OBJ_PROPERTY(application,'activesheet');
	END;
	
	/**********************************************************************************
	// Takes nothing returns nothing
	// -	Releases all handles and destroys the processes related to the opened file
	**********************************************************************************/
	PROCEDURE IMPORT_TERMINATE IS
	BEGIN
			OLE2.RELEASE_OBJ(WORKSHEETS);
			OLE2.RELEASE_OBJ(WORKBOOKS);
			OLE2.RELEASE_OBJ(WORKSHEET);
			OLE2.RELEASE_OBJ(WORKBOOK);
			OLE2.INVOKE(APPLICATION,'QUIT');
			OLE2.RELEASE_OBJ(APPLICATION);
	END;
	
	/**********************************************************************************
	// Takes pathname, block_name, row_start, column_count, success_message, fail message
	//		row_count_tracker, custom_mode returns rows_inserted
	// -	Imports the contents of the target .xlsx file to the target block
	//
	// Parameter specification
	// -  	P_PATHNAME is the destination of the file and filename
	// -	P_BLOCK_NAME is the target block. Use 'DUMMY' instead of ':DUMMY'
	// -	P_ROW_START is the starting row where the data starts. Value is usually set to 1
	//								when there is no header and 2 if there is.
	// -	P_COLUMN_COUNT is the number of rows to be imported
	// - 	P_SUCCESS is the message returned when importing is successful
	// -	P_FAIL is the message returned when importing fail
	// -	ROW_COUNTER if you wanna track the number of affected rows you can use this value
	// -  	P_CUSTOM_MODE when set to true, you must implement IMPORT_CUSTOM_BLOCK to customize
	// 								the fields to fetch and where to fetch them to.
	//
	// Example:
	//		declare
	// 				result varchar2(100);
	//				row_count number
	//		begin
	//				result := CREOXL.IMPORT_TO_BLOCK(
	//											P_PATHNAME => 'c:\demo\path.xlsx', 
	//											P_BLOCK_NAME => 'target_block', 
	//											P_ROW_START => 2. 
	//											P_COLUMN_COUNT => 3, 
	//											P_SUCCESS => 'Import successful', 
	//											P_FAIL => 'Import failed', 
	//											ROW_COUNTER => row_count,
	//											P_CUSTOM_MODE => true
	//									);
	//				message(result || ' Rows imported:' || to_char(row_count));
	//		end;
	//
	// Whats potentially inside 'c:\demo\path.xlsx':
	// P_ROW_START is 2 for this example since we need to skip the first one
	// First Name | Middle Initial | Last Name
	//     Edison |		     B |    Estaca
	//      Clint |              I |  Florento
	//   	  ... |		   ... |       ...
	//--------------------------------------------------------------------------------
	// NOTE: In the event when you need to import a file where you only need a few fields
	// or not all the fields fit into the block in a convenient sequence, you need to set
	// P_CUSTOM_MODE to true when you invoke the IMPORT_TO_BLOCK function and implement the
	// function IMPORT CUSTOM. The code snippet is down below otherwise just return false.
	// IMPORT_TO_BLOCK will call IMPORT_CUSTOM internally as it executes and will pass in
	// the BLOCK_NAME you provided when you invoked it and the current column it is currently
	// working on. You can then check if the column has any significance on what youre working
	// with and then move the cursor to the field where you need its value.
	//
	// SNIPPET:
	// FUNCTION IMPORT_CUSTOM(P_BLOCK_NAME IN VARCHAR2, P_COLUMN_NUMBER IN NUMBER) RETURN BOOLEAN IS
	// BEGIN
	// 		IF P_COLUMN_NUMBER = 1 THEN
	// 				GO_FIELD(P_BLOCK_NAME || '.SKU');
	// 		ELSIF P_COLUMN_NUMBER = 4 THEN
	//	  		GO_FIELD(P_BLOCK_NAME || '.QTY');
	// 		ELSE
	//				RETURN FALSE;
	// 		END IF;				    
	//		RETURN TRUE;
	// END;
	// 
	**********************************************************************************/
	FUNCTION IMPORT_TO_BLOCK(P_PATHNAME IN VARCHAR2, P_BLOCK_NAME IN VARCHAR2, P_ROW_START IN NUMBER, 
		P_COLUMN_COUNT IN NUMBER, P_SUCCESS IN VARCHAR2, P_FAIL IN VARCHAR2, ROW_COUNTER OUT NUMBER, 
		P_CUSTOM_MODE BOOLEAN) RETURN VARCHAR2 IS	
			EOD	BOOLEAN := FALSE;
			EOD_COUNTER	NUMBER;
			CELL_VALUE VARCHAR2(1000);
			FIRST_ENTRY VARCHAR2(1000);
			CURRENT_ITEM VARCHAR2(1000);
			R VARCHAR2(1000);
	BEGIN
			IMPORT_INITIALIZE(P_PATHNAME);
			FIRST_ENTRY := GET_BLOCK_PROPERTY(P_BLOCK_NAME, FIRST_ITEM);
			GO_ITEM(FIRST_ENTRY);
			CLEAR_BLOCK(NO_VALIDATE);
			LAST_RECORD;
			ROW_COUNTER := P_ROW_START;
		 	LOOP EXIT WHEN EOD;
		 			CURRENT_ITEM := FIRST_ENTRY;
		  		EOD_COUNTER := 0;
		  		IF :SYSTEM.record_status <> 'NEW' then
		   				CREATE_RECORD;
		  		END IF;
					FOR COLUMN_COUNTER IN 1..P_COLUMN_COUNT LOOP 
							CELL_VALUE := CREOXL.GET_TEXT(ROW_COUNTER , COLUMN_COUNTER);
		  				IF CELL_VALUE is null THEN
		  	  				EOD_COUNTER := EOD_COUNTER + 1;
		  	  				IF EOD_COUNTER = P_COLUMN_COUNT then
		  	  						GO_FIELD(FIRST_ENTRY);
		  	  						IF to_char(ROW_COUNTER - P_ROW_START) <> 0 then
		  	  								R := P_SUCCESS;
		  	  						ELSIF to_char(ROW_COUNTER - P_ROW_START) = 0 then
		  	  								R := P_FAIL;
		  	  						END IF;
		  								EOD:=true;
		  								EXIT;
		  	  				END IF;
		  				END IF;
		  				IF (P_CUSTOM_MODE) THEN
									IF (IMPORT_CUSTOM_BLOCK(
													P_BLOCK_NAME => P_BLOCK_NAME, 
													P_COLUMN_NUMBER => COLUMN_COUNTER, 
													P_CELL_VALUE => CELL_VALUE)
													
									) THEN
											COPY(CELL_VALUE,NAME_IN('SYSTEM.CURSOR_ITEM'));
									END IF;
							ELSE
									GO_FIELD(P_BLOCK_NAME || '.' || CURRENT_ITEM);
							    COPY(CELL_VALUE,NAME_IN('SYSTEM.CURSOR_ITEM'));
							    CURRENT_ITEM := GET_ITEM_PROPERTY(P_BLOCK_NAME || '.' || CURRENT_ITEM, NEXTITEM);
							END IF;
							
							IF COLUMN_COUNTER = P_COLUMN_COUNT THEN
							  NEXT_RECORD;
							END IF;
					END LOOP; 
			  	ROW_COUNTER := ROW_COUNTER + 1;
			  	SET_RECORD_PROPERTY(:SYSTEM.TRIGGER_RECORD, :SYSTEM.CURSOR_BLOCK, STATUS, NEW_STATUS);
		 	END LOOP;
		 	ROW_COUNTER := ROW_COUNTER - 3;
		 	IMPORT_TERMINATE;
			RETURN R;		
	END;

	/**********************************************************************************
	// Takes pathname, table_name, start_row, column_count and buffer_size returns rows_inserted
	// - inserts the contents of the target excel file into a table and commits it
	//
	// Parameter Specifics:
	// 		P_PATHNAME - path of the target excel file
	//		P_TABLE_NAME - name of the target table
	//		P_ROW_START - starting row where the data starts. Value is usually set to 1
	//				when there is no header and 2 if there is.
	//		P_COLUMN_COUNT is the number of rows to be imported
	//		P_BUFFER_SIZE - rows before reopening the file. 
	// NOTE: When importing from a large file it may fail if you set your buffer too high 
	//				and its going to be slow if its too low. 
	// TIP: When importing from a file with over 20k rows set your buffer to 20k
	//
	// Example:
	// *** IMPORTING TO TABLE ***
	// DECLARE
	// 		row_count number;
	// BEGIN
	// 		row_count := CREOXL.IMPORT_TO_TABLE(
	//							P_PATHNAME => 'c:\example\path\demo.xlsx', 
	//							P_TABLE_NAME => 'table_name', 
	//							P_ROW_START => 2, 
	//							P_COLUMN_COUNT => 4, 
	//							P_BUFFER_SIZE => 20000
	//					  	   );
	// END;
	// 
	// *** IMPORT_CUSTOM_TABLE ***
	//
	// Map the fields to columns and implement some checks if your are unsure about how clean
	// the data would be. If you somehow need to import to different tables add a table name
	// check before checking the column number.
	//
	// PROCEDURE IMPORT_CUSTOM_TABLE(P_TABLE_NAME in varchar2, P_COLUMN_NUMBER in number, P_FIELDS in out varchar2, 
	//	P_VALUES in out varchar2, P_CELL_VALUE in out varchar2) IS
	// BEGIN
	// 	  IF P_COLUMN_NUMBER = 1 THEN
	//	 			P_FIELDS := 'number_column1';
	//	 			P_VALUES := P_CELL_VALUE;
	//		ELSIF P_COLUMN_NUMBER = 1 THEN
	//	 			P_FIELDS := 'text_column2';
	//	 			P_VALUES := '''' || P_CELL_VALUE || '''';
	//		ELSIF P_COLUMN_NUMBER = 4 THEN	
	//				P_FIELDS := 'date_column3';
	//				P_VALUES := '''' || P_CELL_VALUE || '''';
	//		END IF;	
	// END;
	**********************************************************************************/
	FUNCTION IMPORT_TO_TABLE(P_PATHNAME IN VARCHAR2, P_TABLE_NAME IN VARCHAR2, P_ROW_START IN NUMBER, 
			P_COLUMN_COUNT IN NUMBER, P_BUFFER_SIZE IN NUMBER) RETURN NUMBER IS
		PROCESSING BOOLEAN;
		START_ROW NUMBER;
		PROCESSED_ROWS NUMBER := 0;
		PROCESSED_ROWS_CURRENT NUMBER;
		FUNCTION IMPORT_TO_TABLE_CHUNK(P_TABLE_NAME IN VARCHAR2, P_ROW_START IN NUMBER, P_COLUMN_COUNT IN NUMBER, 
					P_MAX_ROWS IN NUMBER) RETURN NUMBER IS	
				EOD	BOOLEAN := FALSE;
				EOD_COUNTER	NUMBER;
				ROW_COUNTER NUMBER;
				CELL_VALUE VARCHAR2(1000);
				TABLE_FIELDS VARCHAR2(10000);
				TABLE_VALUES VARCHAR2(10000);
				TEMP_FIELD VARCHAR2(10000);
				TEMP_VALUE VARCHAR2(10000);
		BEGIN
				ROW_COUNTER := P_ROW_START;
			 	LOOP EXIT WHEN EOD;
			  		EOD_COUNTER := 0;
						TABLE_FIELDS := null;
						TABLE_VALUES := null;
						FOR COLUMN_COUNTER IN 1..P_COLUMN_COUNT LOOP 
								TEMP_FIELD := null;
								TEMP_VALUE := null;
								CELL_VALUE := CREOXL.GET_TEXT(ROW_COUNTER , COLUMN_COUNTER);
								IMPORT_CUSTOM_TABLE(
									P_TABLE_NAME => P_TABLE_NAME,
									P_VALUES => TEMP_VALUE,
									P_FIELDS => TEMP_FIELD,
									P_COLUMN_NUMBER => COLUMN_COUNTER,
									P_CELL_VALUE => CELL_VALUE
								);
								IF (length(TEMP_FIELD) > 0) THEN
									IF (nvl(length(TABLE_FIELDS),0) = 0) THEN
										TABLE_FIELDS := TEMP_FIELD;
									ELSE
										TABLE_FIELDS := TABLE_FIELDS || ',' || TEMP_FIELD;
									END IF;
								END IF;
								IF (length(TEMP_VALUE) > 0) THEN
									IF (nvl(length(TABLE_VALUES),0) = 0) THEN
										TABLE_VALUES := TEMP_VALUE;
									ELSE
										TABLE_VALUES := TABLE_VALUES || ',' || TEMP_VALUE;
									END IF;
								END IF;
								EOD_COUNTER := EOD_COUNTER + 1;
			  				IF CELL_VALUE is null THEN
			  	  				EOD_COUNTER := EOD_COUNTER + 1;
			  	  				IF EOD_COUNTER = P_COLUMN_COUNT then
			  								EOD:=true;
			  								EXIT;
			  	  				END IF;
			  				END IF;
						END LOOP; 
						IF NOT EOD THEN
								FORMS_DDL('INSERT INTO ' || P_TABLE_NAME || ' (' || TABLE_FIELDS || ') VALUES (' || TABLE_VALUES || ')');
								FORMS_DDL('COMMIT');
						END IF;
						IF P_ROW_START + P_BUFFER_SIZE - 1 = ROW_COUNTER THEN
	  					EXIT;
						END IF;
				  	ROW_COUNTER := ROW_COUNTER + 1;
			 	END LOOP;
				IF EOD THEN
						RETURN ROW_COUNTER - P_ROW_START;
				END IF;
			 	RETURN (ROW_COUNTER - P_ROW_START) + 1;
		END;
	BEGIN
		PROCESSING := TRUE;
		START_ROW := P_ROW_START;
		WHILE PROCESSING LOOP
			IMPORT_INITIALIZE(P_PATHNAME);
			PROCESSED_ROWS_CURRENT := IMPORT_TO_TABLE_CHUNK(
				P_TABLE_NAME => P_TABLE_NAME,
				P_ROW_START => START_ROW,
				P_COLUMN_COUNT => P_COLUMN_COUNT,
				P_MAX_ROWS => P_BUFFER_SIZE
			);
			PROCESSED_ROWS := PROCESSED_ROWS + PROCESSED_ROWS_CURRENT;
			START_ROW := START_ROW + P_BUFFER_SIZE;
			IF PROCESSED_ROWS_CURRENT < P_BUFFER_SIZE THEN
				PROCESSING := FALSE;
			END IF;
			IMPORT_TERMINATE;
		END LOOP;
		RETURN PROCESSED_ROWS - 1;
	END;
	
	/**********************************************************************************
	// Takes row and column returns cell_value
	// -	Retrieves the value of the target cell represented by row and column values
	**********************************************************************************/
	FUNCTION GET_TEXT(P_ROW IN NUMBER, P_COLUMN IN NUMBER) RETURN VARCHAR2 IS
		CELL_VALUE varchar2(1000);
	BEGIN
		ARGS:= OLE2.create_arglist;
		OLE2.add_arg(args, P_ROW);
		OLE2.add_arg(args, P_COLUMN);
		CELL:= OLE2.get_obj_property(worksheet, 'Cells', args);
		OLE2.destroy_arglist(args);
		CELL_VALUE :=OLE2.get_char_property(CELL, 'Text');
		CELL_VALUE := rtrim(ltrim(cell_value,'    '),'    '); --these are not space, unknown character
		--CELL_VALUE := rtrim(ltrim(cell_value,' -   '), ' -   ');
		CELL_VALUE := replace(cell_value,'',' ');
		RETURN CELL_VALUE;
	END;
	
END;
