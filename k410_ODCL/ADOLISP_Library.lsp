;;; A library to be used to access a database from Visual LISP
;;; in AutoCAD 2000 or higher using ActiveX Data Objects
;;; (ADO)

;;; Copyright (C) 1999-2006 by The Fleming Group

;;; Permission to use, copy, modify, and distribute this
;;; software for any purpose and without fee is hereby
;;; granted, provided that the above copyright notice
;;; appears in all copies and that both that copyright
;;; notice and the limited warranty and restricted
;;; rights notice below appear in all supporting
;;; documentation.

;;; THE FLEMING GROUP PROVIDES THIS PROGRAM "AS IS" AND WITH
;;; ALL FAULTS. THE FLEMING GROUP SPECIFICALLY DISCLAIMS ANY
;;; IMPLIED WARRANTY OF MERCHANTABILITY OR FITNESS FOR A
;;; PARTICULAR USE.
;;; THE FLEMING GROUP DOES NOT WARRANT THAT THE OPERATION OF
;;; THE PROGRAM WILL BE UNINTERRUPTED OR ERROR FREE.

;;; ------------------------------------------------------------

;;; Revision 2.52 April 2007 by JRF: install workaround for
;;; a bug in AUtoCAD 2008 under Vista.  vl-registry-read of
;;; a REG_EXPAND_SZ value returns (2 . "String") instead of
;;; the correct return value, "String".

;;; Revision 2.51 April 2007 by JRF: Restore 2000i code, with
;;; modifications suggested by Phillippe Absil to maintain
;;; compatibility with KB927779, to fix problem of executing
;;; streod procedures twice when executing an "EXECUTE ..."
;;; statement.

;;; Revision 2.5 March 2007 by JRF: Remove code from
;;; ADOLISP_DoSQL which executed in AutoCAD 2000i and above, and
;;; always execute the code which used to execute only in 2000.
;;; The removed code was incompatible with Microsoft patch
;;; KB927779. Alas, this means that the return value of
;;; ADOLISP_DoSQL is now just T or nil Wwhen the SQL statment is
;;; an INSERT, UPDATE, or DELETE;
;;; the number of rows affected is not available.

;;; Revision 2.4 December 2006 by JRF: Introduced the
;;; ADOLISP_DoNotForceJetODBCParsing global variable, required
;;; to work with Excel files under some circumstances.  Set
;;; it to non-nil to NOT set the Jet OLEDB:ODBC Parsing
;;; property (of the Jet database engine) to true.

;;; Revision 2.32 March 2004 by JRF: Fixed a bug in
;;; ADOLISP_GetTablesAndViews.

;;; Revision 2.31 July 30, 2003 by JRF: Removed
;;; ActualSize from the list of field properties
;;; collected when a SELECT statement is executed:
;;; if no rows were returned (and some other
;;; conditions were true, but I don't know exactly
;;; what conditions) asking about ActualSize
;;; caused an automation error that is untrappable.
;;; It's possible but a litle complex to get
;;; ActualSize back; contact me it you need it.

;;; Revision 2.30 May 1, 2003 by JRF:  Added the
;;; ADOLISP_GetColumns function.

;;; Revision 2.20 April 30, 2003 by JRF: Added the
;;; ADOLISP_FieldsPropertiesList global variable,
;;; containing the properties of the fields
;;; retrieved by the last SQL statement (if it
;;; was a SELECT statement).

;;; Revision 2.15 March 31, 2003 by JRF:
;;; ADOLISP_GetTablesAndViews contained a call to ErrorProcessor
;;; which should be ADOLISP_ErrorProcessor.

;;; Revision 2.14: documentation only.

;;; Revision 2.13 February 3, 2002 by JRF: Fixed a bug in
;;; ADOLISP_GetTablesAndViews which made it always return
;;; (nil nil) when the JET 4.0 driver was being used.

;;; Revision 2.12 May 27, 2002 by JRF: Documentation changes
;;; only, adding information on connecting to Excel to
;;; ADOLISP.DOC.

;;; Revision 2.11 March 14, 2001 by JRF: Fixed bug in
;;; releasing objects after trying to set the properties
;;; of the JET driver in ADOLISP_ConnectToDB

;;; Revision 2.1 March 9, 2002 by JRF: Added code to
;;; ADOLIST_ConnectToDB to set the Jet OLEDB:ODBC Parsing
;;; property to "true" when using the Jet engin, so SQL
;;; statements using double-quotes to surround delimited
;;; identifiers will work.

;; Load the ActiveX stuff for Visual LISP if it isn't already
;; loaded
(vl-load-com)

;;; In case this file gets compiled into a separate-namespace
;;; VLX, export the functions that should be visible.  The
;;; following has no effect unless the document is compiled
;;; into a separate-namespace VLX.
(vl-doc-export 'ADOLISP_ConnectToDB)
(vl-doc-export 'ADOLISP_DoSQL)
(vl-doc-export 'ADOLISP_DisconnectFromDB)
(vl-doc-export 'ADOLISP_ErrorPrinter)
(vl-doc-export 'ADOLISP_GetTablesAndViews)
(vl-doc-export 'ADOLISP_variant-value)

;;; Set up some variables that must be global (within
;;; this file)

;;; Define a VB data type that Visual LISP forgot
(if (not vlax-vbDecimal)
  (setq vlax-vbDecimal 14)
)

;;; Set a flag if we are running in AutoCAD 2000 (not 2000i,
;;; 2002, ...)
(if (< (atof (getvar "ACADVER")) 15.05)
  (setq ADOLISP_IsAutoCAD2000 T)
)

;; Import the ADO type library if it hasn't already been
;; loaded.
(if (null ADOMethod-Append)
  (cond
;; If we can find the library in the registry ...
    ((and (setq ADOLISP_ADODLLPath
                 (vl-registry-read
                   "HKEY_CLASSES_ROOT\\ADODB.Command\\CLSID"
                 )
          )
          (setq ADOLISP_ADODLLPath
                 (vl-registry-read
                   (strcat "HKEY_CLASSES_ROOT\\CLSID\\"
                           ADOLISP_ADODLLPath
                           "\\InProcServer32"
                   )
                 )
          )
          (progn
;; Workaround for bug in AutoCAD 2008 under Vista, returning
;; a dotted pair list containing the string instead of the
;; string
            (if (listp ADOLISP_ADODLLPath)
              (setq ADOLISP_ADODLLPath (cdr ADOLISP_ADODLLPath))
            )
            (findfile ADOLISP_ADODLLPath)
          )
     )
;; Import it
     (vlax-import-type-library
       :tlb-filename ADOLISP_ADODLLPath :methods-prefix
       "ADOMethod-" :properties-prefix "ADOProperty-"
       :constants-prefix "ADOConstant-"
      )
    )
;; Or if we can find it where we expect to find it ...
    ((setq ADOLISP_ADODLLPath
            (findfile
              (if (getenv "systemdrive")
                (strcat
                  (getenv "systemdrive")
                  "\\program files\\common files\\system\\ado\\msado15.dll"
                )
                "c:\\program files\\common files\\system\\ado\\msado15.dll"
              )
            )
     )
;; Import it
     (vlax-import-type-library
       :tlb-filename ADOLISP_ADODLLPath :methods-prefix
       "ADOMethod-" :properties-prefix "ADOProperty-"
       :constants-prefix "ADOConstant-"
      )
    )
;; Can't find the library, tell the user
    (T
     (alert
       (strcat "Cannot find\n\""
               (if ADOLISP_ADODLLPath
                 ADOLISP_ADODLLPath
                 "msado15.dll"
               )
               "\""
       )
     )
    )
  )
)

;;; A routine to connect to a database

;;; Arguments:
;;;     ConnectString:  Either the name of a .UDL file,
;;;                     including the ".UDL", or an
;;;                     OLEDB connection string.
;;;                     If this argument is the name of
;;;                     a UDL file without a full path,
;;;                     it is searched for in the
;;;                     current directory, the
;;;                     AutoCAD search path, and the
;;;                     AutoCAD Data Source Location.
;;;     UserName: The user name to use when connecting.
;;;               May be a null string if the user name is
;;;               specified in the first argument or the
;;;               first argument is a UDL file name.
;;;     Password: The password to use when connecting.
;;                May be a null string if the password is
;;;               supplied in the first argument or the
;;;               first argument is a UDL file name.

;;; Return value:
;;;  If anything fails, NIL.  Call (ADOLISP_ErrorPrinter) to
;;;  print error messages to the command line.
;;;  Otherwise, an ADO Connection Object.

(defun ADOLISP_ConnectToDB (ConnectString UserName Password / IsUDL
                            FullUDLFileName ConnectionObject TempObject
                            ReturnValue ConnectionPropertiesObject
                            ConnectionParsingPropertyObject
                           )
;; Assume no error
  (setq ADOLISP_ErrorList        nil
        ADOLISP_LastSQLStatement nil
  )

;; If the connect string is a UDL file name ...
  (if (= ".UDL"
         (strcase
           (substr ConnectString (- (strlen ConnectString) 3))
         )
      )
    (progn
;; Set a flag that it's a UDL file
      (setq IsUDL T)
;; Try to find it
      (cond
        ((setq FullUDLFileName (findfile ConnectString)))
;; Didn't find it in the current directory or
;; the AutoCAD search path, try the AutoCAD
;; Data Source location
        ((setq FullUDLFileName
                (findfile (strcat (vlax-get-property
                                    (vlax-get-property
                                      (vlax-get-property
                                        (vlax-get-acad-object)
                                        "Preferences"
                                      )
                                      "Files"
                                    )
                                    "WorkspacePath"
                                  )
                                  "\\"
                                  ConnectString
                          )
                )
         )
        )
;; Didn't find it, store an error message
        (t
         (setq ADOLISP_ErrorList
                (list (list (cons "ADOLISP connection error"
                                  (strcat "Can't find \""
                                          ConnectString
                                          "\""
                                  )
                            )
                      )
                )
         )
        )
      )
    )
  )

;; If the first argument is a UDL file name... ...
  (if IsUDL
;; If we found it ...
    (if FullUDLFileName
      (progn
;; Create an ADO connection object
        (setq ConnectionObject
               (vlax-create-object
                 "ADODB.Connection"
               )
        )
;; Try to open the connection.  If there is an error
;; ...
        (if (vl-catch-all-error-p
              (setq TempObject
                     (vl-catch-all-apply
                       'vlax-invoke-method
                       (list ConnectionObject
                             "Open"
                             (strcat "File Name=" FullUDLFileName)
                             UserName
                             Password
                             ADOConstant-adConnectUnspecified
                       )
                     )
              )
            )
          (progn
;; Save the error information
            (setq ADOLISP_ErrorList
                   (ADOLISP_ErrorProcessor TempObject ConnectionObject)
            )
;; Release the connection object
            (vlax-release-object ConnectionObject)
          )
;; It worked, store the connection object in our
;; return value
          (setq ReturnValue ConnectionObject)
        )
      )
    )
;; The connect string is not a UDL file name.
    (progn
;; Create an ADO connection object
      (setq ConnectionObject
             (vlax-create-object "ADODB.Connection")
      )
;; Try to open the connection.  If there is an error ...
      (if (vl-catch-all-error-p
            (setq TempObject
                   (vl-catch-all-apply
                     'vlax-invoke-method
                     (list
                       ConnectionObject "Open" ConnectString UserName
                       Password ADOConstant-adConnectUnspecified
                      )
                   )
            )
          )
        (progn
;; Save the error information
          (setq ADOLISP_ErrorList
                 (ADOLISP_ErrorProcessor TempObject ConnectionObject)
          )
;; Release the connection object
          (vlax-release-object ConnectionObject)
        )
;; It worked, store the connection object in our
;; return value
        (setq ReturnValue ConnectionObject)
      )
    )
  )
;; If we made a connection ...
  (if ReturnValue
    (progn
;; If we want to set ODBC Parsing to true ...
      (if (not ADOLISP_DoNotForceJetODBCParsing)
        (progn
;; Get the properties collection
          (setq ConnectionPropertiesObject
                 (vlax-get-property
                   ReturnValue
                   "Properties"
                 )
          )
;; If the properties collection has a "Jet OLEDB:ODBC
;; Parsing" item ...
          (if (not (vl-catch-all-error-p
                     (setq ConnectionParsingPropertyObject
                            (vl-catch-all-apply
                              'vlax-get-property
                              (list
                                ConnectionPropertiesObject
                                "ITEM"
                                "Jet OLEDB:ODBC Parsing"
                              )
                            )
                     )
                   )
              )
;; Set the "Jet OLEDB:ODBC Parsing" item to
;; "true" so the Jet engine accepts double-quotes
;; around delimited identifiers
            (vlax-put-property
              ConnectionParsingPropertyObject
              "VALUE"
              :vlax-true
            )
          )
        )
      )
;; And release our objects
      (if (= 'VLA-OBJECT (type ConnectionParsingPropertyObject))
        (vlax-release-object ConnectionParsingPropertyObject)
      )
      (if (= 'VLA-OBJECT (type ConnectionPropertiesObject))
        (vlax-release-object ConnectionPropertiesObject)
      )
    )
  )
  ReturnValue
)


;;; A function to execute an arbitrary SQL statement
;;; (replacable parameters are not supported).

;;; Arguments:
;;;     ConnectionObject: An ADO Connection Object.
;;;     SQLString: the SQL statement to execute.

;;; Return value:

;;;  If anything fails, NIL.  Call (ADOLISP_ErrorPrinter) to
;;;  print error messages to the command line.  Otherwise:

;;;  If the SQL statement is a "select ..." statement that
;;;  could return rows, returns a list of lists.  The first
;;;  is a list of the column names.  If any rows were
;;;  returned, the subsequent sub-lists contain the
;;;  returned rows in the same order as the column names
;;;  in the first sub-list.

;;;  If the SQL statement is a "delete ...", "update ...", or
;;;  "insert ..." that cannot return any rows:
;;;    If the program is running in AutoCAD 2000, T
;;;    If the program is running in AutoCAD 2000i or
;;;    later, the integer number of rows affected.

(defun ADOLISP_DoSQL (ConnectionObject SQLStatement /
                      RecordSetObject FieldsObject FieldNumber
                      FieldCount FieldList RecordsAffected
                      TempObject ReturnValue CommandObject
                      IsError FieldItem FieldPropertiesList
                      FieldName
                     )
;; Assume no error
  (setq ADOLISP_ErrorList        nil
;; Initialize global variables
        ADOLISP_LastSQLStatement SQLStatement
        ADOLISP_FieldsPropertiesList nil
  )
;; If we are working in AutoCAD 2000 ...
  (if ADOLISP_IsAutoCAD2000
;; Then we can't use the Execute method of the Command
;; object because returning values in parameters (of a
;; function loaded from an external library) is broken.
    (progn
;; Create an ADO Recordset and set the cursor and lock
;; types
      (setq RecordSetObject
             (vlax-create-object "ADODB.RecordSet")
      )
      (vlax-put-property
        RecordSetObject
        "cursorType"
        ADOConstant-adOpenStatic
      )
      (vlax-put-property
        RecordSetObject
        "LockType"
        ADOConstant-adLockOptimistic
      )
;; Open the recordset.  If there is an error ...
      (if (vl-catch-all-error-p
            (setq TempObject
                   (vl-catch-all-apply
                     'vlax-invoke-method
                     (list RecordSetObject "Open" SQLStatement
                           ConnectionObject nil nil
                           ADOConstant-adCmdText
                          )
                   )
            )
          )
;; Save the error information
        (progn
          (setq ADOLISP_ErrorList
                 (ADOLISP_ErrorProcessor TempObject ConnectionObject)
          )
          (setq IsError T)
          (vlax-release-object RecordSetObject)
        )
;; Otherwise, set an indicator that it worked
        (setq RecordsAffected T)
      )
    )
;; We're in AutoCAD 2000i or above, we can use the
;; Execute method of the Command object and see
;; how many records are affected by an UPDATE, INSERT,
;; or DELETE
    (progn
;; Create an ADO command object and store the query
;; and connection
      (setq CommandObject (vlax-create-object "ADODB.Command"))
      (vlax-put-property
        CommandObject
        "CommandText"
        SQLStatement
      )
      (vlax-put-property
        CommandObject
        "ActiveConnection"
        ConnectionObject
      )
      (vlax-put-property
      CommandObject
      "CommandType"
      ADOConstant-adCmdText
      )

;; Create an ADO Recordset
      (setq RecordSetObject
             (vlax-create-object "ADODB.RecordSet")
      )
;; Open the recordset.  If there is an error ...
      (if (vl-catch-all-error-p
            (setq TempObject
                   (vl-catch-all-apply
                     'vlax-invoke-method
                     (list CommandObject "Execute"
                           nil nil nil
                          )
                   )
            )
          )
;; Save the error information
        (progn
          (setq ADOLISP_ErrorList
                 (ADOLISP_ErrorProcessor TempObject ConnectionObject)
          )
          (setq IsError T)
          (vlax-release-object CommandObject)
          (vlax-release-object RecordSetObject)
        )
        (progn
;; No error, save the recordset
          (setq RecordSetObject TempObject)
        )
      )
    )
  )
;; If there were no errors ...
(if (not IsError)
;; If the recordset is closed ...
    (if (= ADOConstant-adStateClosed
           (vlax-get-property RecordsetObject "State")
        )
;; Then the SQL statement was a "delete ..." or an
;; "insert ..." or an "update ..." which doesn't
;; return any rows.
      (progn
        (setq ReturnValue (not IsError))
;; And release the recordset and command; we're done.
        (vlax-release-object RecordSetObject)
        (if (not ADOLISP_IsAutoCAD2000)
          (vlax-release-object CommandObject)
        )
      )
;; The recordset is open, the SQL statement
;; was a "select ...".
      (progn
;; Get the Fields collection, which
;; contains the names and properties of the
;; selected columns
        (setq FieldsObject (vlax-get-property
                             RecordSetObject
                             "Fields"
                           )
;; Get the number of columns
              FieldCount   (vlax-get-property FieldsObject "Count")
              FieldNumber  -1
        )
;; For all the fields ...
        (while
          (> FieldCount (setq FieldNumber (1+ FieldNumber)))
          (setq FieldItem (vlax-get-property FieldsObject "Item" FieldNumber)
;; Get the names of all the columns in a list to
;; be the first part of the return value
                FieldName (vlax-get-property FieldItem "Name")
                FieldList (cons FieldName FieldList)
                FieldPropertiesList nil
           )
          (foreach FieldProperty '("Type" "Precision" "NumericScale" "DefinedSize" "Attributes")
            (setq FieldPropertiesList (cons (cons FieldProperty (vlax-get-property FieldItem FieldProperty)) FieldPropertiesList))
          )
;; Save the list in the global list
          (setq ADOLISP_FieldsPropertiesList (cons (cons FieldName FieldPropertiesList) ADOLISP_FieldsPropertiesList))
        )
;; Get the FieldsPropertiesList in the right order
        (setq ADOLISP_FieldsPropertiesList (reverse ADOLISP_FieldsPropertiesList))

;; Initialize the return value
        (setq ReturnValue (list (reverse FieldList)))
;; If there are any rows in the recordset ...
        (if
          (not (and (= :vlax-true
                       (vlax-get-property RecordSetObject "BOF")
                    )
                    (= :vlax-true
                       (vlax-get-property RecordSetObject "EOF")
                    )
               )
          )
;; We're about to get tricky, hang on!  Create the
;; final results list ...
           (setq
             ReturnValue
;; By appending the list of rows to the list of
;; fields.
              (append
                (list (reverse FieldList))
;; Uses Douglas Wilson's elegant
;; list-transposing code from
;; http://xarch.tu-graz.ac.at/autocad/lisp/
;; to create the list of rows, because
;; GetRows returns items in column order
                (apply
                  'mapcar
                  (cons
                    'list
;; Set up to convert a list of lists
;; of variants to a list of lists of
;; items that AutoLISP understands
                    (mapcar
                      '(lambda (InputList)
                         (mapcar '(lambda (Item)
                                    (ADOLISP_variant-value Item)
                                  )
                                 InputList
                         )
                       )
;; Get the rows, converting them from
;; a variant to a safearray to a list
                      (vlax-safearray->list
                        (vlax-variant-value
                          (vlax-invoke-method
                            RecordSetObject
                            "GetRows"
                            ADOConstant-adGetRowsRest
                          )
                        )
                      )
                    )
                  )
                )
              )
           )
        )
;; Close the recordset and release it and the
;; command
        (vlax-invoke-method RecordSetObject "Close")
        (vlax-release-object RecordSetObject)
        (if (not ADOLISP_IsAutoCAD2000)
          (vlax-release-object CommandObject)
        )
      )
    )
)
;; And return the results
  ReturnValue
)

;;; A function to close a connection and release
;;; the connection object.

;;; Argument:
;;;    An ADO Connection Object.

;;; Return value:
;;;    Always returns T

(defun ADOLISP_DisconnectFromDB (ConnectionObject)
  (setq ADOLISP_ErrorList        nil
        ADOLISP_LastSQLStatement nil
  )
  (vlax-invoke-method ConnectionObject "Close")
  (vlax-release-object ConnectionObject)
  T
)

;;; ------------------------------------------------------------

;;; ADOLISP utility functions

;;; A function to print the list of errors generated
;;; by the ADOLISP_ErrorProcessor function.  The functions
;;; are separate so ADOLISP_ErrorProcessor can be called
;;; while a DCL dialog box is displayed and then
;;; ADOLISP_ErrorPrinter can be called after the dialog
;;; box has been removed.

;;; No arguments, no return value.

(defun ADOLISP_ErrorPrinter ()
  (if ADOLISP_LastSQLStatement
    (prompt (strcat "\nLast SQL statement:\n\""
                    ADOLISP_LastSQLStatement
                    "\"\n\n"
            )
    )
  )
  (foreach ErrorList ADOLISP_ErrorList
    (prompt "\n")
    (foreach ErrorItem ErrorList
      (prompt
        (strcat (car ErrorItem) "\t\t" (cdr ErrorItem) "\n")
      )
    )
  )
  (prin1)
)

;;; A function to obtain the names of all
;;; the tables and views in a database.
;;; (Views are called "Queries" in Microsoft Access.)

;;; Argument:
;;;     ConnectionObject: An ADO Connection Object

;; Return value:
;;;  A list of two lists.
;;;  The first list contains the table names.
;;;  The second list contains the view names.

(defun ADOLISP_GetTablesAndViews (ConnectionObject / TempObject
                                  TablesList TempList ViewsList
                                 )
  (setq ADOLISP_ErrorList        nil
        ADOLISP_LastSQLStatement nil
  )
  (setq RecordSetObject (vlax-create-object "ADODB.RecordSet"))
;; If we fail getting a recordset of the tables and views
;; ...
  (if (vl-catch-all-error-p
        (setq RecordSetObject
               (vl-catch-all-apply
                 'vlax-invoke-method
                 (list
                   ConnectionObject
                   "OpenSchema"
                   ADOConstant-adSchemaTables
                 )
               )
        )
      )
;; Save the error information
    (setq ADOLISP_ErrorList
           (ADOLISP_ErrorProcessor RecordSetObject ConnectionObject)
    )
    (progn
;; Got the recordset!
;; We're about to get tricky, hang on!  Convert the
;; recordset object to a LISP list ...
      (setq
        TempList
;; Uses Douglas Wilson's elegant
;; list-transposing code from
;; http://xarch.tu-graz.ac.at/autocad/lisp/
;; to create the list of rows, because
;; GetRows returns items in column order
         (apply
           'mapcar
           (cons
             'list
;; Set up to convert a list of lists
;; of variants to a list of lists of
;; items that AutoLISP understands
             (mapcar
               '(lambda (InputList)
                  (mapcar '(lambda (Item)
                             (ADOLISP_variant-value Item)
                           )
                          InputList
                  )
                )
;; Get the rows, converting them from
;; a variant to a safearray to a list
               (vlax-safearray->list
                 (vlax-variant-value
                   (vlax-invoke-method
                     RecordSetObject
                     "GetRows"
                     ADOConstant-adGetRowsRest
                   )
                 )
               )
             )
           )
         )
      )
;; Now filter out the system tables and
;; sort the tables and views into the
;; correct lists
      (foreach Item TempList
        (cond
          ((= (nth 3 Item) "VIEW")
           (setq ViewsList (cons (nth 2 Item) ViewsList))
          )
          ((= (nth 3 Item) "TABLE")
           (setq TablesList (cons (nth 2 Item) TablesList))
          )
        )
      )
;; Close the recordset
      (vlax-invoke-method RecordSetObject "Close")
    )
  )
  (vlax-release-object RecordSetObject)
  (list TablesList ViewsList)
)

;;; A function to obtain the properties
;;; of the columns in a table.

;;; Arguments:
;;;     ConnectionObject: An ADO Connection Object
;;;     TableName: A string containing the table name.
;;;                Not case sensitive.

;;; Return value:
;;;  If nothing was found, NIL.
;;;  If columns were found for that table, a
;;;  list of lists, one sub-list for each column.
;;;  Each sub-list contains:
;;;     Column name
;;;      dotted-pair lists:
;;;         "Type" . OLEDB DataTypeEnum
;;;         "DefinedSize" . Maximum length
;;;                         (character data only)
;;;                         (0 if no maximum)
;;;         "Attributes" . OLEDB FieldAttributeEnum
;;;         "Precision" . number of digits (numerical
;;;                       columns only)
;;;         "Ordinal" . number of the column in the
;;;                     table (the first column is 1)

;;; The sub-lists in the return value will be in
;;; the same order as the ordinal values of the columns.


(defun ADOLISP_GetColumns (ConnectionObject TableName /
                           TempObject TempList ReturnValue
                          )
  (setq ADOLISP_ErrorList
         nil
        ADOLISP_LastSQLStatement
         nil
        TableName (strcase TableName)
  )
  (setq RecordSetObject (vlax-create-object "ADODB.RecordSet"))
;; If we fail getting a recordset of all
;; the columns in the database ...
  (if (vl-catch-all-error-p
        (setq RecordSetObject
               (vl-catch-all-apply
                 'vlax-invoke-method
                 (list
                   ConnectionObject
                   "OpenSchema"
                   ADOConstant-adSchemaColumns
                 )
               )
        )
      )
;; Save the error information
    (setq ADOLISP_ErrorList
           (ADOLISP_ErrorProcessor
             RecordSetObject
             ConnectionObject
           )
    )
    (progn
;; Got the recordset!
;; We're about to get tricky, hang on!  Convert the
;; recordset object to a LISP list ...
      (setq
        TempList
;; Uses Douglas Wilson's elegant
;; list-transposing code from
;; http://xarch.tu-graz.ac.at/autocad/lisp/
;; to create the list of rows, because
;; GetRows returns items in column order
         (apply
           'mapcar
           (cons
             'list
;; Set up to convert a list of lists
;; of variants to a list of lists of
;; items that AutoLISP understands
             (mapcar
               '(lambda (InputList)
                  (mapcar '(lambda (Item)
                             (ADOLISP_variant-value Item)
                           )
                          InputList
                  )
                )
;; Get the rows, converting them from
;; a variant to a safearray to a list
               (vlax-safearray->list
                 (vlax-variant-value
                   (vlax-invoke-method
                     RecordSetObject
                     "GetRows"
                     ADOConstant-adGetRowsRest
                   )
                 )
               )
             )
           )
         )
      )
;; Close the recordset
      (vlax-invoke-method RecordSetObject "Close")
;; Loop over all the columns
      (foreach ColumnList TempList
;; If this column belongs to the correct table ...
        (if (= TableName (strcase (nth 2 ColumnList)))
;; Store its information
          (setq ReturnValue
                 (cons
                   (list (nth 3 ColumnList)
                         (cons "Type" (nth 11 ColumnList))
                         (cons "DefinedSize"
                               (if (nth 13 ColumnList)
                                 (fix (nth 13 ColumnList))
                                 0
                               )
                         )
                         (cons "Attributes"
                               (if (nth 9 ColumnList)
                                 (fix (nth 9 ColumnList))
                                 0
                               )
                         )
                         (cons "Precision"
                               (if (nth 15 ColumnList)
                                 (nth 15 ColumnList)
                                 255
                               )
                         )
                         (cons "Ordinal"
                               (fix (nth 6 ColumnList))
                         )
                   )
                   ReturnValue
                 )
          )
        )
      )
    )
  )
  (vlax-release-object RecordSetObject)

;; The reverse of the return value list is probably in order,
;; but make sure ....
  (if ReturnValue
    (vl-sort (reverse ReturnValue)
             '(lambda (x y)
                (< (cdr (assoc "Ordinal" (cdr x)))
                   (cdr (assoc "Ordinal" (cdr y)))
                )
              )
    )
    nil
  )
)


;;; ------------------------------------------------------------

;;; ADOLISP Support functions

;;; A function to assemble all errors into a list of lists of
;;; dotted pairs of strings ("name" . "value")

(defun ADOLISP_ErrorProcessor (VLErrorObject ConnectionObject /
                               ErrorsObject ErrorObject
                               ErrorCount ErrorNumber ErrorList
                               ErrorValue
                              )
;; First get Visual LISP's error message
  (setq ReturnList   (list
                       (list
                         (cons
                           "Visual LISP message"
                           (vl-catch-all-error-message VLErrorObject)
                         )
                       )
                     )
;; Get the ADO errors object and quantity
        ErrorObject  (vlax-create-object "ADODB.Error")
        ErrorsObject (vlax-get-property ConnectionObject "Errors")
        ErrorCount   (vlax-get-property ErrorsObject "Count")
        ErrorNumber  -1
  )
;; Loop over all the ADO errors ...
  (while (< (setq ErrorNumber (1+ ErrorNumber)) ErrorCount)
;; Get the error object of the current error
    (setq ErrorObject
                      (vlax-get-property ErrorsObject "Item" ErrorNumber)
;; Clear the list of items for this error
          ErrorList   nil
    )
;; Loop over all possible error items of this error
    (foreach ErrorProperty '("Description" "HelpContext"
                             "HelpFile" "NativeError" "Number"
                             "SQLState" "Source"
                            )
;; Get the value of the current item.  If it's a number
;; ...
      (if (numberp (setq ErrorValue
                          (vlax-get-property ErrorObject ErrorProperty)
                   )
          )
;; Convert it to a string for consistency
        (setq ErrorValue (itoa ErrorValue))
      )
;; And store it
      (setq ErrorList (cons (cons ErrorProperty ErrorValue)
                            ErrorList
                      )
      )
    )
;; Add the list for the current error to the return value
    (setq ReturnList (cons (reverse ErrorList) ReturnList))
  )
;; Set up the return value in the correct order
  (reverse ReturnList)
)

;;; A function to convert a variant to a value.  Knows
;;; about more variant types than vlax-variant-value

(defun ADOLISP_variant-value (VariantItem / VariantType)
  (cond
;; If it's a Currency data type or a Decimal data type ...
    ((or (= vlax-vbCurrency
            (setq VariantType (vlax-variant-type VariantItem))
         )
;; Note that I defined vlax-vbDecimal
;; at the beginning of the file
         (= vlax-vbDecimal VariantType)
     )
;; Convert it to a double before getting its value
     (vlax-variant-value
       (vlax-variant-change-type VariantItem vlax-vbDouble)
     )
    )
;; If it's a date, time, or date/time variable type ...
    ((= vlax-vbDate VariantType)
;; Convert it to a string (assuming it's a Microsoft
;; Access type Julian date)
     (1900BasedJulianToCalender
       (vlax-variant-value VariantItem)
     )
    )
;; If it's a boolean value (yes/no, true/false, ...) ...
    ((= vlax-vbBoolean VariantType)
;; Convert it to the string "True" or "False"
     (if (= :vlax-true (vlax-variant-value VariantItem))
       "True"
       "False"
     )
    )
;; If it's an OLE_COLOR data type ...
    ((= vlax-vbOLE_COLOR VariantType)
;; Convert it to a long integer before getting its value
     (vlax-variant-value
       (vlax-variant-change-type VariantItem vlax-vbLong)
     )
    )
;; Otherwise, just turn vlax-variant-value loose on it
    (t (vlax-variant-value VariantItem))
  )
)

;;; A function to convert a "1900-based"Julian-like
;;; date, time, or date/time to a string.

;;; Argument:  A real number, containing a Julian-type date
;;; based on January 1, 1900 (e.g. a Microsoft Access date)
;;; in the integer portion and a time (as a fraction of a
;;; day) in the fractional portion.  Note that this
;;; algorithm considers a number with no fractional
;;; portion to be the day _starting_ at midnight.

;;; Return Value:  A string:
;;;  Containing just the date if there was no fractional
;;;    portion.
;;;  Containing just the time if there was no integer portion
;;;    or the input number was 0.0
;;;  Otherwise, containing the date and the time.

;;; Times are returned as hour:minutes:seconds, 24-hour
;;; format, with leading zeros if necessary to make
;;; two digits per element

;;; Dates are returned in US format (month/day/year) but this
;;; is easily changed.  The year is given as four digits.
;;; The month and day are supplied as two digits (possibly
;;; with leading zeros)

;;; k410 ODCL: Date format changed -- day/month/year

(defun 1900BasedJulianToCalender (JulianDate / a b c d e y z
                                  Month Day Year Hours Minutes
                                  Seconds CalenderTime NoTime
                                  NoDate ReturnValue
                                 )
;; Initialize the return value
  (setq ReturnValue "")
;; If the input date has no time component ...
  (if (equal 0.0
             (float (- JulianDate (float (fix JulianDate))))
             1E-9
      )
;; It has no time component ... if it has no date
;; component ...
    (if (zerop (fix JulianDate))
;; It must be a timestamp of 0:00.00.  Set the flag that
;; we don't have a date but leave the "No Time" flag
;; unset
      (setq NoDate T)
;; It has a date component but has no time component.
;; Shift the date to a real Julian date
      (setq JulianDate (+ 2415019 (fix JulianDate))
;; Set a flag so we know we don't have to
;; calculate the time
            NoTime     T
      )
    )
;; It has a time component.  If it has no date component
;; ...
    (if (zerop (fix JulianDate))
;; Set a flag so we know we don't want to calculate a
;; date
      (setq NoDate T)
;; Otherwise, just shift it to be based like a standard
;; Julian date
      (setq JulianDate (+ 2415019 JulianDate))
    )
  )
;; If we want to calculate the date ...
  (if (not NoDate)
;; It's magic, don't even ask (because I don't know).
;; Some things we weren't meant to know.
    (setq z           (fix JulianDate)
          a           (fix (/ (- z 1867216.25) 36524.25))
          a           (+ z 1 a (- (fix (/ a 4))))
          b           (+ a 1524)
          c           (fix (/ (- b 122.1) 365.25))
          d           (floor (* 365.25 c))
          e           (fix (/ (- b d) 30.6001))
          Day         (fix (- b d (floor (* 30.6001 e))))
          e           (- e
                         (if (< e 14)
                           2
                           14
                         )
                      )
          Month       (1+ e)
          Year        (if (> e 1)
                        (- c 4716)
                        (- c 4715)
                      )
          Year        (if (= Year 0)
                        (1- Year)
                        Year
                      )
;; This uses US format for the date, you might want
;; to change it.

;; k410 ODCL: Date format changed -- day/month/year
          ReturnValue (strcat (if (< Day 10)
                                (strcat "0" (itoa Day))
                                (itoa Day)
                              )
                              "/"
                              (if (< Month 10)
                                (strcat "0" (itoa Month))
                                (itoa Month)
                              )
                              "/"
                              (itoa Year)
                      )
    )
  )
;; If we want to calculate the time ...
  (if (not NoTime)
;; First strip the date portion from the input
    (setq y            (- JulianDate (float (fix JulianDate)))
;; Round to the nearest second
          y            (/ (float (fix (+ 0.5 (* y 86400.0)))) 86400.0)
;; Number of hours since midnight
          Hours        (fix (* y 24))
;; Number of minutes since midnight the hour
;; (1440 minutes per day)
          Minutes      (fix (- (* y 1440.0) (* Hours 60.0)))
;; Number of seconds since the minute (86400
;; seconds per day)
          Seconds      (fix (- (* y 86400.0)
                               (* Hours 3600.0)
                               (* Minutes 60.0)
                            )
                       )
          CalenderTime (strcat (if (< Hours 10)
                                 (strcat "0" (itoa Hours))
                                 (itoa Hours)
                               )
                               ":"
                               (if (< Minutes 10)
                                 (strcat "0" (itoa Minutes))
                                 (itoa Minutes)
                               )
                               ":"
                               (if (< Seconds 10)
                                 (strcat "0" (itoa Seconds))
                                 (itoa Seconds)
                               )
                       )
          ReturnValue  (if (< 0 (strlen ReturnValue))
                         (strcat ReturnValue " " CalenderTime)
                         CalenderTime
                       )

    )
  )
  ReturnValue
)

;;; Floor function, rounds down to the next integer.
;;; Identical with FIX for positive numbers, but
;;; rounds away from zero for negative numbers.

(defun floor (number /)
  (if (> number 0)
    (fix number)
    (fix (- number 1))
  )
)

(prompt "\nADOLISP library loaded\n")