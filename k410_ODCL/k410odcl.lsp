;-------------------------------------------------------------------------------
; k410 ODCL
;-------------------------------------------------------------------------------

;-------------------------------------------------------------------------------
;����������� ���������� �������� � ��������� �������� �����.
;�������� ����� �������� ��������� ��������:
; * K410_ODCL_UDLFILE - ������ ��� (������� ����) ����� UDL ��������� ������;
; * K410_ODCL_TABLENAME - ��� ������� ��������� ������;
; * K410_ODCL_KEYFIELDNAME - ��� ��������� ����;
; * K410_ODCL_ATTRNAME_1 ... K410_ODCL_ATTRNAME_N - ������������ �������� �����
;												���� ������� ��������� ������
;�������� ��������� (Invisible = T), ���������� (Constant = T) � � ��������
;�������� (LockPosition = T).
;
; * K410_ODCL_KEYFIELDVALUE - 	�������� ��������� ���� ������� ��������� ������
;								��� ���������� ���������.
;������� ��������� (Invisible = T) � � �������� �������� (LockPosition = T).
;-------------------------------------------------------------------------------
(setq k410_Ver "\nk410-ODCL Ver.0.06.12.07.20\n")
(princ k410_Ver)
;�������� ���������� AciveX ACAD-a
(vl-load-com)
;�������� ���������� ADOLisp
(load "ADOLISP_Library.lsp")
;�������� �������
(command "OPENDCL")
(dcl_Project_Load "k410odcl" T)

;------------------------------------------------------ ���������� ���������� --
(setq
; - ��������� ����������
		ACADApp (vlax-get-acad-object)
; - ��������� �� �������� ��������
		ActiveDoc (vla-get-ActiveDocument ACADApp)
; - ������ ��� ����� ��������� ������:
		TagUDLFile ""
; - ��� ������� ��������� ������:
		TagTableName ""
; - �������� ���� ������� ��������� ������:
		TagKeyFieldName ""
; - �������� ��������� ���� ������� ��������� ������:
		TagKeyFieldValue ""
; - ������� �������� ��������� ���� ������� ��������� ������:
		AttrRefKeyFieldValue nil
; - ��������� �� ������� �������� �����:
		CurBlock nil
; - ��������� �� ������� ��������� �����:
		CurBlockRef nil
; - ������ ������ �� ���������� �������� �����:
		ConstAttrRefList nil
; - ������ ���������� ��������� �����:
		ConstAttrList nil
; - ������ ������ �� ������� �������� �����:
		AttrRefList nil
; - ������ ������� �������� �����:
		AttrList nil
; - ������ ����� ������� ��������� ������, ��������������� ���������:
		AttrFieldsList nil
; - ������ ������:
		Views nil
; - ������ ����� ������� �������� ��������� ������:
		CurColumnList nil
; - ������� ������ �������� ��������� ����:
		CurKeyValueList nil
; - ������� ����� ������:
		CurRecordSet nil
; - ���� ���������� ������:
		ReloadDataFlag F
; - ���� �������� ����� ���������:		
		NewAttrFlag F
; - ������ �������� ����������:
		ItemIndex 0
; - ���� ������� ����������:
		ReloadFlag 0
; - ������ �������� ������:
		KeyValues nil
; - ������ �������� ������:
		KeyValuesStr ""
; - ����������� ����� ���������� ������ � ������ �������� ������:
		StrDelimiter ", "
; - ���� ������ ������ � ������ ��������:
		KeyValuesClear T
;------------------------------------------------------ ���������� ���������� --

;------------------------------------------------------- ����������-��������� --
		AttrUDLFile "K410_ODCL_UDLFILE"
		AttrTableName "K410_ODCL_TABLENAME"
		AttrKeyFieldName "K410_ODCL_KEYFIELDNAME"
		AttrKeyFieldValue "K410_ODCL_KEYFIELDVALUE"
		AttrPrefix "K410_ODCL_"

		FieldTXT "�������� ����..."
		TableViewTXT "�������� ������� ��� ������..."
		KeyFieldTXT "�������� �������� ����..."
		KeyFieldValueTXT "�������� ��������� ����..."
		ErrorTitleTxt "k410 ODCL - ������"
		TitleTXT "k410 ODCL - "
		
		ArrayWidth "10"
		ArraydX "100"
		ArraydY "50"
		
;�������� �������� ������
		ReloadRecordSet 1
;�������� ������ ���������
		ReloadData 2
;------------------------------------------------------- ����������-��������� --

;------------------------------------------------------------ �������� ������ --
		Error_001_TXT "������ 001: ������ ��������� ����"
;					�������� ����, ����������� � �������� �����	����������� � 
;					������� ��������� ������
		Error_002_TXT "������ 002: ������ �������� UDL ����� ��������� ������"
;					UDL ����� �� ������ �� ������������ ����
		Error_003_TXT "������ 003: ������ �������� ������� ������"
;					� ����������� ������� nil ������ ��������� �� �����, ���
;					������ ������ ����� ������
		Error_004_TXT "������ 004: ������ �������� ��������� ����"
;					��������� � ��������� ���� �������� ��������� ���� 
;					����������� � ������ �������� ��������� ����
;		Error_005_TXT
;------------------------------------------------------------ �������� ������ --
)


;--------------------------------------- ������� ����������� ������ ��������� --
(defun c:k410_Ver (/)
	k410_Ver
)	
;--------------------------------------- ������� ����������� ������ ��������� --

;---------------------------------------------------- ���������� ������� ���� --

;--------------------------- ������� ������ �������� ���� �� ����� (�� Label) --
;����:
; * _Menu - ����, � ������� ������ �����;
; * _LabelString - ��� ������ ����;
;�����:
; * ����� ��������, ���� ����� ����� ����;
; * nil, ���� ������ �������� ���;
(defun FindSimpleLabel (_Menu
						_LabelString /	_MenuItemCount
										_i
										_Result)
;���������� ��������� ����
	(setq	_MenuItemCount (vla-Get-Count _Menu)
;��������� ������� ��������� ����
			_i 0
;��������� ������ �������
			_Result nil)
	(repeat _MenuItemCount
		(if (= (vla-Get-Label (vla-Item _Menu _i)) _LabelString)
			(setq _Result _i))
		(setq _i (1+ _i))
	)
;������� ��������
	_Result
)
;--------------------------- ������� ������ �������� ���� �� ����� (�� Label) --

;-------------------------------------------- ������� ���������� ������� ���� --
(defun AddMenuItems (	_Menu /	_Item_1
                              	_Item_2
								_Item_3)
	(vla-AddSeparator _Menu 0)

	(setq	_Item_3 (vla-AddMenuItem _Menu 0 "������ ���������" "(BlockRefArray) ")
			_Item_2 (vla-AddMenuItem _Menu 0 "����� ������" "(ChoiceRecord) ")
			_Item_1 (vla-AddMenuItem _Menu 0 "������ ������" "(DataAddr) "))

	(if _Item_1 (vla-Put-HelpString _Item_1 "�������� ��� �������������� ������� ������"))
	(if _Item_2 (vla-Put-HelpString _Item_2 "����� ������ ������� ��������� ������"))
	(if _Item_3 (vla-Put-HelpString _Item_3 "�������� ������� ��������� �� ���� ������� ��������� ������"))
	
	(princ)
)
;-------------------------------------------- ������� ���������� ������� ���� --
;�������� ��������� �� PopupMenu
(setq	_MenuGroup (vla-Item (vla-Get-MenuGroups ACADApp) "ACAD")
		_PopupMenu (vla-Item (vla-Get-Menus _MenuGroup) "���� �������� ������ � ����������"))

;���������� ������� ����, ���� ��� �� ���� ���������
(if (FindSimpleLabel _PopupMenu "������ ������")
;then
	(princ)
;else
	(AddMenuItems _PopupMenu)
)
;---------------------------------------------------- ���������� ������� ���� --

;----------------------------------------------------------------- ODCL ����� --

;---------------------------- ������� ����������� ��������� �� ������ ADOLISP --
;������ �������� ������������ ������� ADOLISP_ErrorPrinter ������
;ADOLISP_Library.lsp

(defun k410_ErrorPrinter (/ _ErrorString)
;������������� ����������
	(setq _ErrorString "")
;������������ ������ � ��������� ��������
	(if ADOLISP_LastSQLStatement
		(setq _ErrorString (strcat	"Last SQL statement:\n\""
									ADOLISP_LastSQLStatement
									"\"\n\n")
		)
	)
;���������� ���������� �� ������
	(foreach ErrorList ADOLISP_ErrorList
		(foreach ErrorItem ErrorList
			(setq _ErrorString
				(strcat _ErrorString (car ErrorItem) ": " (cdr ErrorItem) "\n")
			)
		)
	)
;����������� ��������� � ���������� ����
	(dcl_MessageBox _ErrorString ErrorTitleTxt 2 4)
;"�����" �����
	(princ)
)
;---------------------------- ������� ����������� ��������� �� ������ ADOLISP --

;------------------------------- ������� �������� ������ �� ��������� � ����� --
;�������� ������������ ������� ADOLISP_DoSQL ������ ADOLISP_Library.lsp
;��������� ��������� ������ ������ SELECT ... ������ Command �� ���������
;����:
; - ConnectionObject - ����������,
; - SQLStatement - ����� �������.
;�����:
; - nil - � ������ ������� ��� ���������� ������ ������;
; - ������ - ((������ ����� ������ ������) vla-object �����������)

(defun k410_DoSQL (	ConnectionObject 
					SQLStatement	/	RecordSetObject
										ReturnValue
										FieldsObject
										FieldList
										FieldNumber
										FieldCount
										FieldItem
										FieldName
										FieldPropertiesList
										TempObject
										IsError)
;����� ����������
;; Assume no error
	(setq ADOLISP_ErrorList nil
;; Initialize global variables
        ADOLISP_LastSQLStatement SQLStatement
        ADOLISP_FieldsPropertiesList nil)
;; Create an ADO Recordset and set the cursor and lock
;; types
		(setq RecordSetObject (vlax-create-object "ADODB.RecordSet"))
		(vlax-put-property RecordSetObject "CursorType" ADOConstant-adOpenStatic)
		(vlax-put-property RecordSetObject "LockType" ADOConstant-adLockOptimistic)
		(vlax-put-property RecordSetObject "CursorLocation" ADOConstant-adUseClient)
		(vlax-put-property RecordSetObject "Source" SQLStatement)
;; Open the recordset.  If there is an error ...
		(if (vl-catch-all-error-p
				(setq TempObject 	(vl-catch-all-apply 'vlax-invoke-method
										(list 	RecordSetObject "Open" 
												SQLStatement
												ConnectionObject nil nil
												ADOConstant-adCmdText
										)
									)
				)
			)
;; Save the error information
				(progn
					(setq ADOLISP_ErrorList (ADOLISP_ErrorProcessor TempObject ConnectionObject))
					(setq IsError T)
					(vlax-release-object RecordSetObject)
				)
		)
;; If there were no errors ...
	(if (not IsError)
;; If the recordset is closed ...				
		(if (= ADOConstant-adStateClosed (vlax-get-property RecordsetObject "State"))
;; Then the SQL statement was a "delete ..." or an		
;; "insert ..." or an "update ..." which doesn't		
;; return any rows.						
			(progn
				(setq ReturnValue (not IsError))
;; And release the recordset and command; we're done.		
				(vlax-release-object RecordSetObject)
				(if (not ADOLISP_IsAutoCAD2000) (vlax-release-object CommandObject))
			)
;; The recordset is open, the SQL statement			
;; was a "select ...".						
			(progn
;; Get the Fields collection, which				
;; contains the names and properties of the			
;; selected columns						
				(setq	FieldsObject (vlax-get-property RecordSetObject "Fields")
;; Get the number of columns					
						FieldCount   (vlax-get-property FieldsObject "Count")
						FieldNumber  -1)
;; For all the fields ...					
				(while (> FieldCount (setq FieldNumber (1+ FieldNumber)))
					(setq FieldItem (vlax-get-property FieldsObject "Item" FieldNumber)
;; Get the names of all the columns in a list to		
;; be the first part of the return value			
						FieldName (vlax-get-property FieldItem "Name")
						FieldList (cons FieldName FieldList)
						FieldPropertiesList nil)
					(foreach FieldProperty '("Type" "Precision" "NumericScale" "DefinedSize" "Attributes")
						(setq FieldPropertiesList (cons (cons FieldProperty (vlax-get-property FieldItem FieldProperty)) FieldPropertiesList)))
;; Save the list in the global list				
					(setq ADOLISP_FieldsPropertiesList (cons (cons FieldName FieldPropertiesList) ADOLISP_FieldsPropertiesList))
				)

;; Get the FieldsPropertiesList in the right order
				(setq ADOLISP_FieldsPropertiesList (reverse ADOLISP_FieldsPropertiesList))
;; Initialize the return value
				(setq ReturnValue (list (reverse FieldList) RecordSetObject))		
			)
		)
	)
;; And return the results
	ReturnValue
)
;------------------------------- ������� �������� ������ �� ��������� � ����� --

;------------------------------------------- ������� ������ �������� � ������ --

;----------------------------- ������� �������� ��������� �� �������� ����� c --
;							   �������� ���������� ��������� 
;����:
; *	_BlockRef - ������ �� ��������� �����;
; *	_AttrProperty - ������ � ������ ����������� ��������:
;	- Backward;
;	- Constant;
;	- HasExtensionDictionary;
;	- Invisible;
;	- LockPosition;
;	- MTextAttribute;
;	- Preset;
;	- UpsideDown;
;	- Verify;
; *	_EqFlag - �������� ���������� (:vlax-true ��� :vlax-false);
;�����:
; *	������ ������ �� �������� �� �������� �����
; *	nil � ������ ���������� ��������� � �����
(setq 	AttrObjectName 		"AcDbAttributeDefinition")

(defun GetAttributes (	_BlockRef 
						_AttrProperty 
						_EqFlag /	_TempVar
									_Block
									_Index
									_IntermediateResult
									_Result)
;����� � ��������� ����������
	(setq 	_Index 0
			_Result nil
			_Block (vla-Item (vla-get-Blocks (vla-get-ActiveDocument (vlax-get-acad-object))) (vla-get-EffectiveName _BlockRef)))
;����� �������� "�������" �� ��������� "�������"
	(repeat (vla-get-Count _Block)
		(setq _TempVar (vla-Item _Block _Index))
		(if (= (vla-get-ObjectName _TempVar) AttrObjectName)
			(progn
				(if (= (strcase _AttrProperty) "BACKWARD") (setq _IntermediateResult (vla-get-Backward _TempVar)))
				(if (= (strcase _AttrProperty) "CONSTANT") (setq _IntermediateResult (vla-get-Constant _TempVar)))
				(if (= (strcase _AttrProperty) "HASEXTENSIONDICTIONARY") (setq _IntermediateResult (vla-get-HasExtensionDictionary _TempVar)))
				(if (= (strcase _AttrProperty) "INVISIBLE") (setq _IntermediateResult (vla-get-Invisible _TempVar)))
				(if (= (strcase _AttrProperty) "LOCKPOSITION") (setq _IntermediateResult (vla-get-LockPosition _TempVar)))
				(if (= (strcase _AttrProperty) "MTEXTATTRIBUTE") (setq _IntermediateResult (vla-get-MTextAttribute _TempVar)))
				(if (= (strcase _AttrProperty) "PRESET") (setq _IntermediateResult (vla-get-Preset _TempVar)))
				(if (= (strcase _AttrProperty) "UPSIDEDOWN") (setq _IntermediateResult (vla-get-UpsideDown _TempVar)))
				(if (= (strcase _AttrProperty) "VERIFY") (setq _IntermediateResult (vla-get-Verify _TempVar)))
				(if (= _IntermediateResult _EqFlag) (setq _Result (append _Result (list _TempVar))))
			)
		)
		(setq _Index (1+ _Index))
	)
  _Result
)
;							   ������� �������� ��������� �� �������� ����� c
;----------------------------- �������� ���������� ��������� -------------------

;------------------ ������� ������������� ���������� ��� ���������� ��������� --
(defun InitBlock (/	_Index
					_TempVar)
;��������� ��������� �� ��������� ���������
	(if (not ReloadDataFlag)
		(setq CurBlockRef (vla-Item (vla-get-PickfirstSelectionSet ActiveDoc) 0))
	)
;��������� ��������� �� �������� �����
	(setq 	CurBlock (vla-Item (vla-get-Blocks ActiveDoc) (vla-get-EffectiveName CurBlockRef))
;�������� ������ ������ �� ������� �������� �����
			AttrRefList (GetAttributes CurBlockRef "Constant" :vlax-false)
;�������� ������ ������ �� ����������� �������� �����
			ConstAttrRefList (GetAttributes CurBlockRef "Constant" :vlax-true)
;����� ����������
			ConstAttrList nil
			AttrList nil
			AttrFieldsList nil
			AttrRefKeyFieldValue nil
			TagUDLFile ""
			TagTableName ""
			TagKeyFieldName ""
			TagKeyFieldValue ""
			NewAttrFlag F)
;������������ ������ ���� ����������� ���������
	(foreach _ConstAttr ConstAttrRefList
		(setq ConstAttrList (append ConstAttrList (list (vla-get-TagString _ConstAttr))))
;������ �������� UDLFile
		(if (= (vla-get-TagString _ConstAttr) AttrUDLFile)
			(setq TagUDLFile (vla-get-TextString _ConstAttr));then
		)
;������ �������� ����� ������� ��������� ������	
		(if (= (vla-get-TagString _ConstAttr) AttrTableName)
			(setq TagTableName (vla-get-TextString _ConstAttr));then
		)
;������ ����� ��������� ���� ������� ��������� ������		
		(if (= (vla-get-TagString _ConstAttr) AttrKeyFieldName)
			(setq TagKeyFieldName (vla-get-TextString _ConstAttr));then
		)		
	)
;������������ ������ ���� ������� ���������
	(foreach _CurAttr AttrRefList
		(if (/= (vla-get-TagString _CurAttr) AttrKeyFieldValue)
;then
			(progn
				(setq AttrList (append AttrList (list (vla-get-TagString _CurAttr))))
				(if (null (vl-position (strcat AttrPrefix (vla-get-TagString _CurAttr)) ConstAttrList))
;then
					(setq AttrFieldsList (append AttrFieldsList (list "")))
;else
					(setq AttrFieldsList 
						(append AttrFieldsList 
							(list	
								(vla-get-TextString 
									(nth (vl-position (strcat AttrPrefix (vla-get-TagString _CurAttr)) ConstAttrList) ConstAttrRefList)
								)
							)
						)
					)
				)
			)
		)
	)
;����� �������� �������� ��������� ���� ������� ��������� ������
	(foreach _AttrRef (vlax-safearray->list (vlax-variant-value (vla-GetAttributes CurBlockRef)))
;�������� �������� ��������� ����
		(if (= (vla-Get-TagString _AttrRef) AttrKeyFieldValue)
;then
			(setq 	TagKeyFieldValue (vla-Get-TextString _AttrRef)
					AttrRefKeyFieldValue _AttrRef)
		)
	)
)
;------------------ ������� ������������� ���������� ��� ���������� ��������� --

;----------------------------------------- ������� ������������ ������ ������ --
;�����:
; - CurColumnList - ������ ���� ����� ������������ ������ ������
; - CurRecordSet - ������ �� ����������� ����� ������

(defun GetRecordSet (/	_KeyFields
						_KeyValues
						_FullView
						_ConnectionObj)

;����� ������������ ������ ������
	(setq 	_View (assoc (strcat TagUDLFile ";" TagTableName ";" TagKeyFieldName) Views)
;����� ����������
			CurRecordSet nil
			CurColumnList nil
			CurKeyValueList nil
			;ADOLISP_DoNotForceJetODBCParsing nil
			)
	(if (null _View)
;then
;����������� ��������� ������
		(if (/= TagUDLFile "")
;then
			(if (findfile TagUDLFile)
;����������� � ��������� ������
				(progn
					(prompt "�������� ���������� � ���������� ������\n")
					(setq _ConnectionObj (ADOLISP_ConnectToDB TagUDLFile "" ""))
					(if (null _ConnectionObj)
;then
						(k410_ErrorPrinter)
;else
						(progn
							(prompt "�������� ������ ������\n")
							(setq	_KeyFields (last (k410_DoSQL _ConnectionObj (strcat "SELECT " TagKeyFieldName " FROM " TagTableName ";")))
									_FullView (k410_DoSQL _ConnectionObj (strcat "SELECT * FROM " TagTableName ";")))
							(prompt "�������� �������� ��������� ����\n")
;���������������� ������� �� ������ ������ ������ ������
							(vlax-invoke-method _KeyFields "MoveFirst")
;�������� �������� ��������� ����
							(while (= (vlax-get-property _KeyFields "EOF") :vlax-false)
								(setq _KeyValues 	
									(append	_KeyValues
										(list
											(vl-string-right-trim "\r"
												(vlax-invoke-method _KeyFields "GetString" 2 1 nil nil nil)
											)
										)
									)
								)
							)
;�������� ������ ������ �������� ��������� ����
							(vlax-invoke-method _KeyFields "Close")
							(vlax-release-object _KeyFields)
;������������ ��������� ������							
							(setq	_View (list (car _FullView) 
												_KeyValues
												(last _FullView)))													
;�������� �� ���������� nil-�� � ���������� �������
							(if (null (vl-position nil _View))
;then								
								(progn
;������ ����� ������� ��������� ������
									(setq 	CurColumnList (car _View)
;������ �������� ��������� ����
											CurKeyValueList _KeyValues
;����� ������
											CurRecordSet (last _FullView)
;���������� ������ ������ � ������ �������
											Views (append 	Views 
															(list (append (list (strcat TagUDLFile ";" TagTableName ";" TagKeyFieldName)) _View))))
;�������� ��������� ����
									(if (null (vl-position TagKeyFieldName CurColumnList))
										(alert Error_001_TXT))
								)
;else
								(alert Error_003_TXT)
							)
						)
					)
				)
;else						
				(alert Error_002_TXT)					
			)
		)
;else
;������ ����� ������� ��������� ������
		(setq 	CurColumnList (cadr _View)
;������ �������� ��������� ����
				CurKeyValueList (caddr _View)
;����� ������
				CurRecordSet (last _View))
	)
)
;----------------------------------------- ������� ������������ ������ ������ --

;------------------------------------------- ������� ���������� ������ ������ --
(defun ReloadRecordSet (/)
;��������� ����� ������� ����������
	(setq ReloadFlag ReloadRecordSet)
;������ ���������� ������ ������
	(prompt "\n")
;����������� �����
	(dcl_Form_Show k410odcl_ReloadDialog)
;����� ����� ������� ����������
	(setq ReloadFlag 0)
)
;------------------------------------------- ������� ���������� ������ ������ --

;------------------------------------------ ������� ���������� �������� ����� --
;�����:
; - nil - �������� ����� ��� ��� �������� ����� ����� nil
; - ������ �������� KeyValues � � ������ �������� KeyValuesStr (���������� 
;	������	������ ������������ ��� ��������� SQL-�������� � ���������� ������)
;
;���� ���������� ���� KeyValuesClear (= �) ����� ����������� ������� ������
;�������� ������������

;������� ��������� ������� �������� ������ �� �������� ���������
(defun GetKeyValues (/	_Index
							_PFSS
							_Item)
;������������� ����������
	(setq 	_Index 0
			ReloadDataFlag T
			_PFSS (vla-get-PickFirstSelectionSet ActiveDoc))
;����� ������� �������� ������
	(if (= KeyValuesClear T)
		(setq 	KeyValues nil
				KeyValuesStr "")
	)
;���������� �������� ��������� ����
	(repeat  (vla-get-Count _PFSS)
;�������� ������ �� ������� ���������
		(setq _Item (vla-Item _PFSS _Index))
;��������� �������� ��������� ����
		(if (= (vla-get-ObjectName _Item) "AcDbBlockReference")
			(if (= (vla-get-HasAttributes _Item) :vlax-true)
				(progn
					(setq CurBlockRef _Item)
;������������� ���������� �����
					(InitBlock)
					(if (/= TagKeyFieldValue "")
						(progn
;���������� �������� ����� � ������
							(setq KeyValues (append KeyValues (list TagKeyFieldValue)))
;���������� �������� ����� � ������ ��������
							(if	(= KeyValuesStr "")
								(setq KeyValuesStr TagKeyFieldValue);then
								(setq KeyValuesStr (strcat KeyValuesStr StrDelimiter TagKeyFieldValue));else
							)
						)
					)
				)
			)
		)
		(setq _Index (1+ _Index))
	)
;����� ����� ���������� ������
	(setq ReloadDataFlag F)
)
;������� �������
(defun c:k410_GetKeyValues (/)
	(GetKeyValues)
)
;�������� �������
;����� ����������� ������� ������ � ������ �������� ������������
(defun c:k410_gkv (/)
	(GetKeyValues)
;������� ������������ ������
	KeyValuesStr
)
;-------------------------------------------------------------------------------
;����� ������ �������� ������ KeyValues � KeyValuesStr
(defun c:k410_KeyValuesClear (/)
;����� ������ � ������ ��������
	(setq 	KeyValues nil
			KeyValuesStr "")
)
;------------------------------------------ ������� ���������� �������� ����� --

;----------------------------------------------------- ������� OR ��� ������� --
;����:
;	- _L1, _L2 - ������
;�����:
;	������, ��������� �� ��������� _L1 � ��������� _L2
(defun ListOR (_L1 _L2 /)
	(if (= (and (listp _L1) (listp _L2)) T)
;����������� ������� � ����������� ����������
		(append _L1 _L2);then
		(prompt "\n����������� ������� ������ ���� ������\n");else
	)
)
;----------------------------------------------------- ������� OR ��� ������� --

;---------------------------------------------------- ������� XOR ��� ������� --
;����:
;	- _L1, _L2 - ������
;�����:
;	������, ��������� �� ��������� _L1, ������� ��� � _L2 � ��������� _L2,
;	������� ��� � _L1
(defun ListXOR (	_L1 _L2 /	_Item
								_FullList
								_Part1
								_Part2)
	(if (= (and (listp _L1) (listp _L2)) T)
		(progn
;C�����, ��������� �� ���� ��������� �������
			(setq 	_FullList (append _L1 _L2)
					_Part1 _FullList
					_Part2 _FullList)
;�������� ��������� ������� ������
			(foreach _Item _L1
				(setq _Part1 (vl-remove _Item _Part1))
			)
;�������� ��������� ������� ������
			(foreach _Item _L2
				(setq _Part2 (vl-remove _Item _Part2))
			)
;����������� ���������� ������ � ����������� ����������
			(append _Part2 _Part1)
		)
		(prompt "\n����������� ������� ������ ���� ������\n")
	)
)
;---------------------------------------------------- ������� XOR ��� ������� --

;---------------------------------------------------- ������� AND ��� ������� --
;����:
;	- _L1, _L2 - ������
;�����:
;	������, ��������� �� ���������, ������� ���� � �_L1 � � _L2
(defun ListAND (	_L1 _L2 /	_Item
								_FullList
								_Part)
	(if (= (and (listp _L1) (listp _L2)) T)
		(progn
;C�����, ��������� �� ���� ��������� �������
			(setq 	_FullList (append _L1 _L2)
					_Part _FullList)
;�������� ��������� ������� ������ 
;(��������� ������ ��������� _L1, ������� ��� � _L2)
			(foreach _Item _L2
				(setq _FullList (vl-remove _Item _FullList))
			)
;�������� ���������, ������� ��� �� ������ ������
			(foreach _Item _FullList
				(setq _L1 (vl-remove _Item _L1))
			)
;����������� ����������
			_L1
		)
		(prompt "\n����������� ������� ������ ���� ������\n")
	)
)
;---------------------------------------------------- ������� AND ��� ������� --

;###################################################### ������: ������ ������ ##
;------------------------- ������� �������� ��� �������������� ������� ������ --
(defun DataAddr (/)
	(prompt "\n")
	(if (= (vla-get-ActiveSpace ActiveDoc) 1)
		(dcl_Form_Show k410odcl_DataAddrDialog);then
		(alert "������� ����������� ������ ��� �������� ������������ ������");else
	)
	(princ)
;��������, ���� ������� ����� ��������
	(if NewAttrFlag 
		(progn
			(command "_.battman"))
			(setq NewAttrFlag F)
		)
)
;------------------------- ������� �������� ��� �������������� ������� ������ --

;------------------------------------------------------ ������������� ������� --
(defun c:k410odcl_DataAddrDialog_OnInitialize (/ _Index)
;������������� ����������
	(InitBlock)
;����� ��������� � ������ ����������
	(dcl_Control_SetProperty k410odcl_DataAddrDialog_FieldNameLabel "Enabled" F)
	(dcl_Control_SetProperty k410odcl_DataAddrDialog_FieldComboBox "Enabled" F)
	(dcl_Control_SetProperty k410odcl_DataAddrDialog_FieldComboBox "Text" FieldTXT)
	(dcl_Control_SetProperty k410odcl_DataAddrDialog_AddRecordButton "Enabled" F)
	(dcl_Control_SetProperty k410odcl_DataAddrDialog_DelRecordButton "Enabled" F)
	(dcl_Control_SetProperty k410odcl_DataAddrDialog_SaveButton "Enabled" F)
	(dcl_Grid_Clear k410odcl_DataAddrDialog_Grid)
	(dcl_Control_SetProperty k410odcl_DataAddrDialog_LabelUDLFilePrompt "Caption" TagUDLFile)
	(dcl_Control_SetProperty k410odcl_DataAddrDialog_LabelTableNamePrompt "Caption" TagTableName)
	(dcl_Control_SetProperty k410odcl_DataAddrDialog_LabelKeyFieldPrompt "Caption" TagKeyFieldName)
;����� ����������
	(setq _Index 0)
;������ ������ ���� ������� ���������
	(repeat (length AttrList)
		(dcl_Grid_AddRow k410odcl_DataAddrDialog_Grid (nth _Index AttrList) (nth _Index AttrFieldsList))
		(if (/= (nth _Index AttrFieldsList) "") (dcl_Control_SetProperty k410odcl_DataAddrDialog_DelRecordButton "Enabled" T))
		(setq _Index (1+ _Index))
	)
;�������� ������ ������
	(GetRecordSet)
	(if (not (null CurColumnList))
		(progn
			(dcl_Control_SetProperty k410odcl_DataAddrDialog_FieldNameLabel "Enabled" T)
			(dcl_Control_SetProperty k410odcl_DataAddrDialog_FieldComboBox "Enabled" T)
			(dcl_Control_SetProperty k410odcl_DataAddrDialog_FieldComboBox "List" CurColumnList)
		)
	) 
)
;------------------------------------------------------ ������������� ������� --

;----------------------------------------------------- ������� ��������� ���� --
(defun c:k410odcl_DataAddrDialog_FieldComboBox_OnSelChanged (ItemIndexOrCount Value /)
;��������� ������ "�������� ������"
	(dcl_Control_SetProperty k410odcl_DataAddrDialog_AddRecordButton "Enabled" T)
)
;----------------------------------------------------- ������� ��������� ���� --

;--------------------------------------------------- ������ "�������� ������" --
(defun c:k410odcl_DataAddrDialog_DSNameButton_OnClicked (/ _Index)
;�������� ������� �� ������ ������ ��������� ������
	(dcl_Form_Show k410odcl_OpenUDLForm)
;����� ��������� ���������� � ������� ������� ������, ���� ��������� ��������
;������
	(if (= _DSChangeFlag T)
;then
		(progn
			(dcl_Control_SetProperty k410odcl_DataAddrDialog_AddRecordButton "Enabled" F)
			(dcl_Control_SetProperty k410odcl_DataAddrDialog_FieldComboBox "Text" FieldTXT)

			(if (= _DSClearAddr T)
				(progn
					(dcl_Control_SetProperty k410odcl_DataAddrDialog_DelRecordButton "Enabled" F)
					(setq _Index 0)
						(repeat (length (dcl_Grid_GetColumnCells k410odcl_DataAddrDialog_Grid 0))
							(dcl_Grid_SetCellText k410odcl_DataAddrDialog_Grid _Index 1 "")
							(setq _Index (1+ _Index))
						)
				)
			)
		)
	)  
)
;--------------------------------------------------- ������ "�������� ������" --

;--------------------------------------------------- ������ "�������� ������" --
(defun c:k410odcl_DataAddrDialog_AddRecordButton_OnClicked (/	_Attr
																_Field
																_Row)
;�������� ����� ����
	(setq 	_Field (dcl_Control_GetProperty k410odcl_DataAddrDialog_FieldComboBox "Text")
;��������� ������ ������
			_Row (car (dcl_Grid_GetCurCell k410odcl_DataAddrDialog_Grid)))
;���������� ������ � �������
	(if (/= _Row -1)
;then
		(progn 
			(dcl_Grid_SetCellText k410odcl_DataAddrDialog_Grid _Row 1 _Field)
;����������� ������ "������� ������"
   
			(dcl_Control_SetProperty k410odcl_DataAddrDialog_DelRecordButton "Enabled" T)
		)
	)
;����������� ������ "���������"
	(dcl_Control_SetProperty k410odcl_DataAddrDialog_SaveButton "Enabled" T)
)
;--------------------------------------------------- ������ "�������� ������" --

;---------------------------------------------------- ������ "������� ������" --
(defun c:k410odcl_DataAddrDialog_DelRecordButton_OnClicked (/	_Row
																_�ellList)
;��������� ������ ������
	(setq _Row (car (dcl_Grid_GetCurCell k410odcl_DataAddrDialog_Grid)))
;������� ������
	(dcl_Grid_SetCellText k410odcl_DataAddrDialog_Grid _Row 1 "")
;��������� ������ "���������"
	(dcl_Control_SetProperty k410odcl_DataAddrDialog_SaveButton "Enabled" T)
;�������� �� ������ ������ � ���������� ������, ���� ��� ������ �������
	(setq _�ellList (dcl_Grid_GetColumnCells k410odcl_DataAddrDialog_Grid 1))
	(dcl_Control_SetProperty k410odcl_DataAddrDialog_DelRecordButton "Enabled" F)
	(repeat (length _�ellList)
		(if (/= (car _�ellList) "")
;then
			(dcl_Control_SetProperty k410odcl_DataAddrDialog_DelRecordButton "Enabled" T)
		)
		(setq _�ellList (cdr _�ellList))
	)
;����������� ������ "���������"
	(dcl_Control_SetProperty k410odcl_DataAddrDialog_SaveButton "Enabled" T)  
)
;---------------------------------------------------- ������ "������� ������" --

;--------------------------------------------------------- ������ "���������" --
(defun c:k410odcl_DataAddrDialog_SaveButton_OnClicked (/	_AttrIndex
															_Index
															_AttrList
															_FieldList
															_AttrName)
;����� �������� UDLFile
	(setq _AttrIndex (vl-position AttrUDLFile ConstAttrList))
;������ �������� UDLFile
	(if (null _AttrIndex)
;then
;�������� �������� UDL �����
		(setq 	_AttrIndex 
				(vla-AddAttribute	CurBlock
									0.001																				;Height
									(+ 	acAttributeModeInvisible														;Mode
									acAttributeModeConstant
									acAttributeModeLockPosition)
									"k410 ODCL: ��������� �������"														;Prompt
									(vlax-safearray-fill (vlax-make-safearray vlax-vbDouble '(0 . 2)) '(0.0 0.0 0.0))	;Insertion point
									AttrUdlFile																			;Tag
									TagUDLFile																			;Value
				)
				ConstAttrList (append ConstAttrList (list AttrUdlFile))
				ConstAttrRefList (append ConstAttrRefList (list _AttrIndex))
;��������� ����� �������� ������ ��������
				NewAttrFlag T
		)
;else
;������ � ������� UDL �����
		(vla-put-textstring (nth _AttrIndex ConstAttrRefList) TagUDLFile)
	)
;����� �������� ������� ��������� ������
	(setq _AttrIndex (vl-position AttrTableName ConstAttrList))
;������ �������� ����� ������� ��������� ������
	(if (null _AttrIndex)
;then
;�������� �������� ����� ������� ��������� ������
		(setq 	_AttrIndex 
				(vla-AddAttribute	CurBlock
									0.001																				;Height
									(+ 	acAttributeModeInvisible														;Mode
										acAttributeModeConstant
										acAttributeModeLockPosition)
									"k410 ODCL: ��������� �������"														;Prompt
									(vlax-safearray-fill (vlax-make-safearray vlax-vbDouble '(0 . 2)) '(0.0 0.0 0.0))	;Insertion point
									AttrTableName																		;Tag
									TagTableName																		;Value
				)	 
				ConstAttrList (append ConstAttrList (list AttrTableName))
				ConstAttrRefList (append ConstAttrRefList (list _AttrIndex))
;��������� ����� �������� ������ ��������
				NewAttrFlag T
		)
;else
;������ � ������� ����� ������� ��������� ������
		(vla-put-textstring (nth _AttrIndex ConstAttrRefList) TagTableName)
	)
;����� �������� ��������� ���� ������� ��������� ������
	(setq _AttrIndex (vl-position AttrKeyFieldName ConstAttrList))
;������ ����� ��������� ���� ������� ��������� ������
	(if (null _AttrIndex)
;then
;�������� �������� ��������� ���� ������� ��������� ������
		(setq 	_AttrIndex 
				(vla-AddAttribute	CurBlock
									0.001																				;Height
									(+ 	acAttributeModeInvisible														;Mode
										acAttributeModeConstant
										acAttributeModeLockPosition)
									"k410 ODCL: ��������� �������"														;Prompt
									(vlax-safearray-fill (vlax-make-safearray vlax-vbDouble '(0 . 2)) '(0.0 0.0 0.0))	;Insertion point
									AttrKeyFieldName																	;Tag
									TagKeyFieldName																		;Value
				)
				ConstAttrList (append ConstAttrList (list AttrKeyFieldName))
				ConstAttrRefList (append ConstAttrRefList (list _AttrIndex))
;��������� ����� �������� ������ ��������
				NewAttrFlag T
		)
;else
;������ � ������� ��������� ���� ������� ��������� ������
		(vla-put-textstring (nth _AttrIndex ConstAttrRefList) TagKeyFieldName)
	)
;������ �������� ��������� ���� ������� ��������� ������
	(if (null AttrRefKeyFieldValue)
;then
;�������� �������� ��������� ���� ������� ��������� ������
		(setq 	_AttrIndex 
				(vla-AddAttribute	CurBlock
									0.001																				;Height
									(+ 	acAttributeModeInvisible														;Mode
										acAttributeModeLockPosition)
									"k410 ODCL: ��������� �������"														;Prompt
									(vlax-safearray-fill (vlax-make-safearray vlax-vbDouble '(0 . 2)) '(0.0 0.0 0.0))	;Insertion point
									AttrKeyFieldValue																	;Tag
									TagKeyFieldValue																	;Value
				)			
				AttrRefKeyFieldValue _AttrIndex
;��������� ����� �������� ������ ��������
				NewAttrFlag T
		)
;else
;������ � ������� ��������� ���� ������� ��������� ������
		(vla-put-textstring AttrRefKeyFieldValue TagKeyFieldValue)
	)
;������������� ���������
	(setq 	_Index 0
			_AttrList (dcl_Grid_GetColumnCells k410odcl_DataAddrDialog_Grid 0)
			_FieldList (dcl_Grid_GetColumnCells k410odcl_DataAddrDialog_Grid 1)
;������ ������� ������
			_Index 0)
	(repeat (length _AttrList)
		(setq 	_AttrName (strcat AttrPrefix (nth _Index _AttrList))
				_AttrIndex (vl-position _AttrName ConstAttrList))
		(if (null _AttrIndex)
;then
;�������� �������� ��������� ���� ������� ��������� ������
			(setq 	_AttrIndex 
					(vla-AddAttribute	CurBlock
										0.001																				;Height
										(+ 	acAttributeModeInvisible														;Mode
											acAttributeModeConstant
											acAttributeModeLockPosition)
										"k410 ODCL: ��������� �������"														;Prompt
										(vlax-safearray-fill (vlax-make-safearray vlax-vbDouble '(0 . 2)) '(0.0 0.0 0.0))	;Insertion point
										_AttrName																			;Tag
										(nth _Index _FieldList)																;Value
					)
					ConstAttrList (append ConstAttrList (list _AttrName))
					ConstAttrRefList (append ConstAttrRefList (list _AttrIndex))
;��������� ����� �������� ������ ��������
					NewAttrFlag T
			)
;else
;������ � ������� �������� ���� ������� ��������� ������
			(vla-put-textstring (nth _AttrIndex ConstAttrRefList) (nth _Index _FieldList))
		)
		(setq _Index (1+ _Index))
	)
;���������� ������ "���������
	(dcl_Control_SetProperty k410odcl_DataAddrDialog_SaveButton "Enabled" F)  
)
;--------------------------------------------------------- ������ "���������" --

;----------------------------------------------------------- ������ "�������" --
(defun c:k410odcl_DataAddrDialog_CancelButton_OnClicked (/)
	(dcl_Form_Close k410odcl_DataAddrDialog 2)
)
;----------------------------------------------------------- ������ "�������" --

;------------------ �������� ������� "������� ���� �������� ��������� ������" --
(defun c:k410odcl_OpenUDLForm_OnClose (	UpperLeftX
										UpperLeftY /	_ConnectionObj
														_View)
;�������� ������� ����� (������� ����) ����� UDL ��������� ������
	(setq TagUDLFile (dcl_FileExplorer_GetPathName k410odcl_OpenUDLForm_FileExplorer))
;�������� ����� ����� UDL ����� ��������� ������
	(if (> (strlen TagUDLFile) 255)
		(progn
			(dcl_MessageBox	"����� ������� ����� ����� UDL ��������� ������ (������� ����)\n �� ������ ��������� 255 ��������."
							ErrorTitleTxt
							2 4)
			(setq TagUDLFile "")
		)  
	)
;����� ����� ������ ������ ��������� ������
	(setq _DSChangeFlag F)
;����������� ��������� ������
	(if (/= TagUDLFile "")
;then
;����������� � ��������� ������
		(progn
			(prompt "����������� � ��������� ������\n")
			(setq _ConnectionObj (ADOLISP_ConnectToDB TagUDLFile "" ""))
			(if (null _ConnectionObj)
;then
				(k410_ErrorPrinter)
;else
				(progn
;����������� ������� ������ ��������� ������
					(dcl_Form_Show k410odcl_DSChoiceDialog)
;�������� ����������
					(ADOLISP_DisconnectFromDB _ConnectionObj)
				)  
			)	
		)
;else
		(princ)
	)
)
;------------------ �������� ������� "������� ���� �������� ��������� ������" --

;---------------------- ������������� ������� ������ ������� � ��������� ���� --
;������ �����������, ���� ���������� ���� �������
(defun c:k410odcl_DSChoiceDialog_OnInitialize (/	_TV
													_TablesAndViews)
;�������� ������� ������ � ��������

	(setq 	_TV (ADOLISP_GetTablesAndViews _ConnectionObj)
			_TablesAndViews (append (nth 0 _TV) (nth 1 _TV)))
;������ ������� ������ � �������� � ���������� ������
	(dcl_Control_SetProperty k410odcl_DSChoiceDialog_DSTableComboBox "List" _TablesAndViews)
;����� ��������� ����� ���������� �������, ������
	(dcl_Control_SetProperty k410odcl_DSChoiceDialog_DSTableComboBox "Text" TableViewTXT)
	(dcl_Control_SetProperty k410odcl_DSChoiceDialog_DSKeyFieldComboBox "Text" KeyFieldTXT)
	(dcl_Control_SetProperty k410odcl_DSChoiceDialog_DSKeyFieldComboBox "Enabled" F)
	(dcl_Control_SetProperty k410odcl_DSChoiceDialog_DSKeyFieldLabel "Enabled" F)
	(dcl_Control_SetProperty k410odcl_DSChoiceDialog_OkButton "Enabled" F)
)
;---------------------- ������������� ������� ������ ������� � ��������� ���� --

;------------------------------------------------ ������� ����� ����� ������� --
(defun c:k410odcl_DSChoiceDialog_DSTableComboBox_OnSelChanged (	ItemIndexOrCount 
																Value / _CL)
;����� ���������� ���� ����������� ������
	(dcl_Control_SetProperty k410odcl_DSChoiceDialog_DSKeyFieldComboBox "Text" KeyFieldTXT)
;��������� ������ ����� ������� ��� �������
	(setq _CL (ADOLISP_GetColumns _ConnectionObj Value))
	(if (null _CL)
;then
		(progn
;��������� �� ������
			(k410_ErrorPrinter)
;����� ����������� ������
			(dcl_Control_SetProperty k410odcl_DSChoiceDialog_DSKeyFieldComboBox "Enabled" F)
			(dcl_Control_SetProperty k410odcl_DSChoiceDialog_DSKeyFieldLabel "Enabled" F)       
		)
;else		
		(progn
;��������� ����������� ������
			(dcl_Control_SetProperty k410odcl_DSChoiceDialog_DSKeyFieldComboBox "Enabled" T)
			(dcl_Control_SetProperty k410odcl_DSChoiceDialog_DSKeyFieldLabel "Enabled" T)
;������������ � ������ ����������� ������
			(setq CurColumnList nil)
			(foreach _ColName _CL
				(setq 	CurColumnList (append CurColumnList (list (car _ColName))))
			)
			(dcl_Control_SetProperty k410odcl_DSChoiceDialog_DSKeyFieldComboBox "List" CurColumnList)
		)
	)
)
;------------------------------------------------ ������� ����� ����� ������� --

;----------------------------------------------- ������� ����� ��������� ���� --
(defun c:k410odcl_DSChoiceDialog_DSKeyFieldComboBox_OnSelChanged (ItemIndexOrCount Value /)
;����������� ������ "��"
	(dcl_Control_SetProperty k410odcl_DSChoiceDialog_OkButton "Enabled" T)
)
;----------------------------------------------- ������� ����� ��������� ���� --

;------------- ������ "Ok" � �������� ������� ������ ������� � ��������� ���� --
(defun c:k410odcl_DSChoiceDialog_OkButton_OnClicked (/	_View
														_Col
														_Str
														_Prop)
;���������� ����� ������� � ��������� ����
	(setq 	TagTableName (dcl_Control_GetProperty k410odcl_DSChoiceDialog_DSTableComboBox "Text")
			TagKeyFieldName (dcl_Control_GetProperty k410odcl_DSChoiceDialog_DSKeyFieldComboBox "Text"))
;������ ������� ����� � ���������� ������
	(dcl_Control_SetProperty k410odcl_DataAddrDialog_FieldComboBox "List" CurColumnList)
	(dcl_Control_SetProperty k410odcl_DataAddrDialog_FieldNameLabel "Enabled" T)
	(dcl_Control_SetProperty k410odcl_DataAddrDialog_FieldComboBox "Enabled" T)
	(dcl_Control_SetProperty k410odcl_DataAddrDialog_SaveButton "Enabled" T)
	(dcl_Control_SetProperty k410odcl_DataAddrDialog_LabelUDLFilePrompt "Caption" TagUDLFile)
	(dcl_Control_SetProperty k410odcl_DataAddrDialog_LabelTableNamePrompt "Caption" TagTableName)
	(dcl_Control_SetProperty k410odcl_DataAddrDialog_LabelKeyFieldPrompt "Caption" TagKeyFieldName)  
;�������� ������� � ����� 1 ([Enter])
	(dcl_Form_Close k410odcl_DSChoiceDialog 1)
;��������� ����� ������ ������ ��������� ������
	(setq _DSChangeFlag T)
;��������� ����� ������ ������� ������
	(if (= (dcl_Control_GetProperty k410odcl_DSChoiceDialog_CheckBox "Value") 1)
;then
		(setq _DSClearAddr F)
;else
		(setq _DSClearAddr T)
	)
)
;------------- ������ "Ok" � �������� ������� ������ ������� � ��������� ���� --
;###################################################### ������: ������ ������ ##

;####################################################### ������: ����� ������ ##
(defun ChoiceRecord (/)
	(prompt "\n")
	(if (= (vla-get-ActiveSpace ActiveDoc) 1)
		(dcl_Form_Show k410odcl_ChoiceRecordDialog);then
		(alert "������� ����������� ������ ��� �������� ������������ ������");else
	)
	(princ)
)


;--------------------------------------------------------- ���������� ������� --
(defun FillChoiceRecordDialogTable (/	_Str
										_KeyPosition)
;����� ��� ��������� ������ "�������� ����� ������"
	(if (null CurRecordSet)
		(dcl_Control_SetProperty k410odcl_ChoiceRecordDialog_ReloadDSButton "Enabled" F)
		(dcl_Control_SetProperty k410odcl_ChoiceRecordDialog_ReloadDSButton "Enabled" T)
	)	
;�������� ������� ����������� ������ (���� ��� �����������)	
	(while (/= (dcl_Grid_GetColumnCount k410odcl_ChoiceRecordDialog_Grid) 0)
		(dcl_Grid_DeleteColumn k410odcl_ChoiceRecordDialog_Grid 0)
	)
;�������� ������� �������
	(if (not (null CurColumnList))
		(foreach _ColName (reverse CurColumnList)
			(dcl_Grid_InsertColumn k410odcl_ChoiceRecordDialog_Grid 0 _ColName)
		)
	)
	(if (= TagKeyFieldValue "")
;then
		(dcl_Control_SetProperty k410odcl_ChoiceRecordDialog_KeyFieldTextBox "Text" KeyFieldValueTXT)
;else
		(progn
			(dcl_Control_SetProperty k410odcl_ChoiceRecordDialog_KeyFieldTextBox "Text" TagKeyFieldValue)
;���������� ������� ��������� ���� � ������ �������� ��������� ����
			(setq _KeyPosition (vl-position TagKeyFieldValue CurKeyValueList))
			(if (null _KeyPosition)
;then
;����� ��������� �� ������
				(alert Error_004_TXT)
;else
				(progn
;����������� ������� �� ����������� �������
					(vlax-invoke-method CurRecordSet "Move" _KeyPosition 1)
;�������� ������
					(setq 	_Str (vlax-invoke-method CurRecordSet "GetString" 2 1 "\t" "\t" nil))
;������ � ������ ������������� ������������	"\t\t" �� "\t \t"
					(repeat (length CurColumnList) (setq _Str (vl-string-subst "\t \t" "\t\t" _Str)))
;������� ������� �� ���������� ��������
					(dcl_Grid_Clear k410odcl_ChoiceRecordDialog_Grid)
;������ � ������� ����������� ������
					(dcl_Grid_AddString k410odcl_ChoiceRecordDialog_Grid _Str)
;��������� ������ "Ok"			
					(dcl_Control_SetProperty k410odcl_ChoiceRecordDialog_OkButton "Enabled" T)
				)
			)
		)
	)
)
;--------------------------------------------------------- ���������� ������� --

;---------------------------------------- ������������� ������� ������ ������ --
(defun c:k410odcl_ChoiceRecordDialog_OnInitialize (/)
;������������� ����������	
	(InitBlock)
;������ �������� ������������� ��������� ������
	(dcl_Control_SetProperty k410odcl_ChoiceRecordDialog_LabelUDLFilePrompt "Caption" TagUDLFile)
	(dcl_Control_SetProperty k410odcl_ChoiceRecordDialog_LabelTableNamePrompt "Caption" TagTableName)
	(dcl_Control_SetProperty k410odcl_ChoiceRecordDialog_LabelKeyFieldPrompt "Caption" TagKeyFieldName)
;����� ��������� ����������
	(dcl_Control_SetProperty k410odcl_ChoiceRecordDialog_SearchButton "Enabled" F)
	(dcl_Control_SetProperty k410odcl_ChoiceRecordDialog_OkButton "Enabled" F)
;�������� ������ ������
	(GetRecordSet)
;���������� ������� ������� ������ ������
	(FillChoiceRecordDialogTable)
;����� �������� ����������
	(setq ItemIndex 0)
;��������� ������ �� ���� �����	
	(dcl_Control_SetFocus k410odcl_ChoiceRecordDialog_KeyFieldTextBox)
)
;---------------------------------------- ������������� ������� ������ ������ --

;---------------------------------------------- ����� �������� ��������� ���� --
(defun c:k410odcl_ChoiceRecordDialog_KeyFieldTextBox_OnEditChanged (NewValue /)
;��������� ������ "�����"
	(dcl_Control_SetProperty k410odcl_ChoiceRecordDialog_SearchButton "Enabled" T)
;����� ������ "Ok"			
	(dcl_Control_SetProperty k410odcl_ChoiceRecordDialog_OkButton "Enabled" F) 
)
;---------------------------------------------- ����� �������� ��������� ���� --

;----------------------------------------------------------- ������ "�������" --
(defun c:k410odcl_ChoiceRecordDialog_CancelButton_OnClicked (/)
	(dcl_Form_Close k410odcl_ChoiceRecordDialog)
)
;----------------------------------------------------------- ������ "�������" --

;--------------------------------------------------------------- ����� ������ --
(defun RecordSearch (/ 	_KeyPosition
						_Str)
;���������� ������� ��������� ���� � ������ �������� ��������� ����
	(setq _KeyPosition (vl-position (dcl_Control_GetProperty k410odcl_ChoiceRecordDialog_KeyFieldTextBox "Text") CurKeyValueList))
	(if (null _KeyPosition)
;then
;����� ��������� �� ������
		(alert Error_004_TXT)
;else
		(progn
;����������� ������� �� ����������� �������
			(vlax-invoke-method CurRecordSet "Move" _KeyPosition 1)
;�������� ������
			(setq 	_Str (vlax-invoke-method CurRecordSet "GetString" 2 1 "\t" "\t" nil))
;������ � ������ ������������� ������������	"\t\t" �� "\t \t"
			(repeat (length CurColumnList) (setq _Str (vl-string-subst "\t \t" "\t\t" _Str)))
;������� ������� �� ���������� ��������
			(dcl_Grid_Clear k410odcl_ChoiceRecordDialog_Grid)
;������ � ������� ����������� ������
			(dcl_Grid_AddString k410odcl_ChoiceRecordDialog_Grid _Str)
;����� ������ "�����"
			(dcl_Control_SetProperty k410odcl_ChoiceRecordDialog_SearchButton "Enabled" F)
;��������� ������ "Ok"			
			(dcl_Control_SetProperty k410odcl_ChoiceRecordDialog_OkButton "Enabled" T)
		)
	)
)
;--------------------------------------------------------------- ����� ������ --

;------------------------------------------------------------- ������ "�����" --
(defun c:k410odcl_ChoiceRecordDialog_SearchButton_OnClicked (/)
	(RecordSearch)
)
;------------------------------------------------------------- ������ "�����" --

;------------------------------------------------- ������ �������� � �������� --
(defun WriteToAttrs (/	_AttrIndex
						_ColPos
						_Layers)
;���� ���������
	(setq _Layers (vla-get-Layers ActiveDoc))
;������ �������� � �������� �� �������
	(foreach _AttrRef (vlax-safearray->list (vlax-variant-value (vla-GetAttributes CurBlockRef)))
;����� �������� ��������
		(setq _AttrIndex (vl-position (vla-Get-TagString _AttrRef) AttrList))
;������ �������� � �������
		(if (not (null _AttrIndex))
;then
			(progn
;����������� �������� � ������ �������
				(setq _ColPos (vl-position (nth _AttrIndex AttrFieldsList) CurColumnList))
				(if (= (vla-get-Lock (vla-Item _Layers (vla-get-Layer _AttrRef))) :vlax-true)
					(prompt "������� �� ������������� ����\n")	;then
;������ �������� � �������
					(progn	;else
						(if (not (null _ColPos))
;then				
							(vla-Put-TextString _AttrRef (dcl_Grid_GetCellText k410odcl_ChoiceRecordDialog_Grid 0 _ColPos))
;else
							(if (= (dcl_Control_GetProperty k410odcl_ChoiceRecordDialog_CheckBox2 "Value") 1)
								(vla-Put-TextString _AttrRef "")
							)
						)
;������ � ������� �������� ��������� ����
						(if (not ReloadDataFlag)
;then
							(vla-Put-TextString AttrRefKeyFieldValue (dcl_Control_GetProperty k410odcl_ChoiceRecordDialog_KeyFieldTextBox "Text"))
						)
					)
				)
			)
		)
	)
)
;------------------------------------------------- ������ �������� � �������� --

;---------------------------------------------------------------- ������ "Ok" --
(defun c:k410odcl_ChoiceRecordDialog_OkButton_OnClicked (/)
	(WriteToAttrs)
;���������� ������
	(dcl_Control_SetProperty k410odcl_ChoiceRecordDialog_OkButton "Enabled" F)
;�������� �������
	(if (= (dcl_Control_GetProperty k410odcl_ChoiceRecordDialog_CheckBox1 "Value") 1)
		(dcl_Form_Close k410odcl_ChoiceRecordDialog)
	)
)
;---------------------------------------------------------------- ������ "Ok" --

;--------------------------------------------- ������ "�������� ����� ������" --
(defun c:k410odcl_ChoiceRecodDialog_ReloadDSButton_OnClicked (/)
;��������� ����� ������
	(ReloadRecordSet)
;���������� ������� ������� ������ ������
	(FillChoiceRecordDialogTable)
)
;--------------------------------------------- ������ "�������� ����� ������" --

;--------------------------------------------------- ������ "�������� ������" --
(defun c:k410odcl_ChoiceRecordDialog_ReloadDataButton_OnClicked (/)
;��������� ����� ������� ����������
	(setq ReloadFlag ReloadData)
;������ ���������� ������ ������
	(prompt "\n")
;����������� �����
	(dcl_Form_Show k410odcl_ReloadDialog)
;����� ����� ������� ����������
	(setq ReloadFlag 0)
)
;--------------------------------------------------- ������ "�������� ������" --

;----------------------------------- CheckBox "��������� �������� ��� ������" --
(defun c:k410odcl_ChoiceRecordDialog_FillErrorCheckBox_OnClicked (Value /)
	(if (= Value 1)
		(progn	;then
			(dcl_Control_SetProperty k410odcl_ChoiceRecordDialog_FillErrorLabel "Enabled" T)
			(dcl_Control_SetProperty k410odcl_ChoiceRecordDialog_FillErrorTextBox "Enabled" T)
		)
		(progn	;else
			(dcl_Control_SetProperty k410odcl_ChoiceRecordDialog_FillErrorLabel "Enabled" F)
			(dcl_Control_SetProperty k410odcl_ChoiceRecordDialog_FillErrorTextBox "Enabled" F)
		)
	)
)
;----------------------------------- CheckBox "��������� �������� ��� ������" --

;------------------------------------------------------------- ������ "Enter" --
(defun c:k410odcl_ChoiceRecordDialog_OnOK (/)
  (RecordSearch)
  (WriteToAttrs)
)
;------------------------------------------------------------- ������ "Enter" -- 
;####################################################### ������: ����� ������ ##

;################################################### ������: ������ ��������� ##
(defun BlockRefArray (/)
	(prompt "\n")
	(if (= (vla-get-ActiveSpace ActiveDoc) 1)
		(dcl_Form_Show k410odcl_BlockRefArray);then
		(alert "������� ����������� ������ ��� �������� ������������ ������");else
	)
	(princ)
)


;------------------------- ������������� ������� ���������� ������� ��������� --
(defun c:k410odcl_BlockRefArray_OnInitialize (/)
;���������� ����� ��������� ����������
	(if (= (dcl_Control_GetProperty k410odcl_BlockRefArray_WidthTextBox "Text") "")
		(dcl_Control_SetProperty k410odcl_BlockRefArray_WidthTextBox "Text" ArrayWidth))
		
	(if (= (dcl_Control_GetProperty k410odcl_BlockRefArray_dYTextBox "Text") "")
		(dcl_Control_SetProperty k410odcl_BlockRefArray_dYTextBox "Text" ArraydX))
		
	(if (= (dcl_Control_GetProperty k410odcl_BlockRefArray_dXTextBox "Text") "")
		(dcl_Control_SetProperty k410odcl_BlockRefArray_dXTextBox "Text" ArraydY))

;����� ��������� ����������
	(dcl_Control_SetProperty k410odcl_BlockRefArray_WidthTextBox "Enabled" F)
	(dcl_Control_SetProperty k410odcl_BlockRefArray_dYTextBox "Enabled" F)
	(dcl_Control_SetProperty k410odcl_BlockRefArray_dXTextBox "Enabled" F)
	(dcl_Control_SetProperty k410odcl_BlockRefArray_OkButton "Enabled" F)
	(dcl_Control_SetProperty k410odcl_BlockRefArray_ReloadDSButton "Enabled" F)
	(if (= KeyValues nil)
		(progn
			(dcl_Control_SetProperty k410odcl_BlockRefArray_OptionButton1 "Value" 1)
			(dcl_Control_SetProperty k410odcl_BlockRefArray_OptionButton2 "Enabled" F)
		)
		(dcl_Control_SetProperty k410odcl_BlockRefArray_OptionButton2 "Enabled" T)
	)
;������������� �����
	(InitBlock)
;�������� ������ ������
	(GetRecordSet)
;��������� ��������� ����������
	(if (not (null CurKeyValueList))
		(progn
			(dcl_Control_SetProperty k410odcl_BlockRefArray_WidthTextBox "Enabled" T)
			(dcl_Control_SetProperty k410odcl_BlockRefArray_dYTextBox "Enabled" T)
			(dcl_Control_SetProperty k410odcl_BlockRefArray_dXTextBox "Enabled" T)
			(dcl_Control_SetProperty k410odcl_BlockRefArray_OkButton "Enabled" T)
			(dcl_Control_SetProperty k410odcl_BlockRefArray_ReloadDSButton "Enabled" T)
		)
	)
)

;------------------------- ������������� ������� ���������� ������� ��������� --

;------------------------------------------------- �������� ������� ��������� --
(defun CreateBlockRefArray (/	_KeyValueList
								_MSpace
								_Width
								_dX
								_dY
								_InsPoint
								_BlockName
								_BasePoint
								_Index)
;�������� ����������� ���������� �������
	(setq	_MSpace (vla-get-ModelSpace ActiveDoc)
			_Width (atoi (dcl_Control_GetProperty k410odcl_BlockRefArray_WidthTextBox "Text"))
			_dX (atoi (dcl_Control_GetProperty k410odcl_BlockRefArray_dXTextBox "Text"))
			_dY (atoi (dcl_Control_GetProperty k410odcl_BlockRefArray_dYTextBox "Text"))
			_BlockName (vla-get-EffectiveName CurBlockRef)
			ReloadDataFlag T
			_Index 0)
;�������� ������ �������� ������
	(if (= (dcl_Control_GetProperty k410odcl_BlockRefArray_OptionButton1 "Value") 1)
		(setq _KeyValueList CurKeyValueList);then
		(setq _KeyValueList KeyValues);else
	)
;�������� ���� �������
	(dcl_Form_Close k410odcl_BlockRefArray)
;����� ������� ����� �������
	(setq 	_InsPoint (getpoint)
			_BasePoint _InsPoint)
;���������� �������
	(foreach _KeyFieldValue _KeyValueList
;���������� ���������
		(setq 	CurBlockRef (vla-InsertBlock _MSpace (vlax-3d-point _InsPoint) _BlockName 1 1 1 0)
				_Index (1+ _Index))
;������������� ���������� ���������
		(InitBlock)
;������ �������� ��������� ���� �� ���������
		(vla-Put-TextString AttrRefKeyFieldValue _KeyFieldValue)
;������ ����� ������� ���������� ���������
		(if (= _Index _Width)
			(setq	_Index 0
					_InsPoint (list (car _BasePoint) (+ (cadr _InsPoint) _dY) 0))
			(setq _InsPoint (list (+ (car _InsPoint) _dX) (cadr _InsPoint) 0))
		)
	)
	(setq ReloadDataFlag F)
)
;------------------------------------------------- �������� ������� ��������� --

;---------------------------------------------------------------- ������ "Ok" --
(defun c:k410odcl_BlockRefArray_OkButton_OnClicked (/)
	(CreateBlockRefArray)
)
;---------------------------------------------------------------- ������ "Ok" --

;----------------------------------------------------------- ������ "�������" --
(defun c:k410odcl_BlockRefArray_CancelButton_OnClicked (/)
  (dcl_Form_Close k410odcl_BlockRefArray)
)

;----------------------------------------------------------- ������ "�������" --

;--------------------------------------------- ������ "�������� ����� ������" --
(defun c:k410odcl_BlockRefArray_ReloadDSButton_OnClicked (/)
 ;��������� ����� ������
	(ReloadRecordSet)
)
;--------------------------------------------- ������ "�������� ����� ������" --

;------------------------------------------------------------- ������ "Enter" --
(defun c:k410odcl_BlockRefArray_OnOK (/)
  (CreateBlockRefArray))
;------------------------------------------------------------- ������ "Enter" --

;------------------------------------------------- �������� �������� �������� --
(defun c:k410odcl_BlockRefArray_WidthTextBox_OnEditChanged (NewValue /)
	(if (and	(<= (atoi NewValue) 0)
					(/= NewValue ""))
		(progn
			(alert "���������� ����� ��� �������� ������ ���� ������ 0")
			(dcl_Control_SetProperty k410odcl_BlockRefArray_WidthTextBox "Text" ArrayWidth)
		)
	)
)

(defun c:k410odcl_BlockRefArray_WidthTextBox_OnKillFocus (/)
	(if (= (dcl_Control_GetProperty k410odcl_BlockRefArray_WidthTextBox "Text") "")
		(progn
			(alert "���������� ����� ��� �������� ������ ���� ������ 0")
			(dcl_Control_SetProperty k410odcl_BlockRefArray_WidthTextBox "Text" ArrayWidth)
		)
	)
)
;-------------------------------------------------------------------------------
(defun c:k410odcl_BlockRefArray_dYTextBox_OnKillFocus (/)
	(if (= (dcl_Control_GetProperty k410odcl_BlockRefArray_dYTextBox "Text") "")
		(progn
			(alert "�������� ���������� ����� ������")
			(dcl_Control_SetProperty k410odcl_BlockRefArray_dYTextBox "Text" ArraydY)
		)
	)
)
;-------------------------------------------------------------------------------
(defun c:k410odcl_BlockRefArray_dXTextBox_OnKillFocus (/)
	(if (= (dcl_Control_GetProperty k410odcl_BlockRefArray_dXTextBox "Text") "")
		(progn
			(alert "�������� ���������� ����� ���������")
			(dcl_Control_SetProperty k410odcl_BlockRefArray_dXTextBox "Text" ArraydX)
		)
	)
)

;------------------------------------------------- �������� �������� �������� --
;################################################### ������: ������ ��������� ##

;######################################################### ������: ���������� ##
;------------------------------------------------------ ������������� ������� --
(defun c:k410odcl_ReloadDialog_OnInitialize (/)
;���������� ������ ������
	(if (= ReloadFlag ReloadRecordSet)
		(progn
;��������� ����
			(dcl_Control_SetProperty k410odcl_ReloadDialog "TitleBarText" (strcat TitleTXT "�������� ����� ������"))
;����������� ������ ��������� � ������
			(dcl_Control_SetProperty k410odcl_ReloadDialog_ProgressBar "Value" 0)
;����������� ������ �������
			(dcl_Control_SetProperty k410odcl_ReloadDialog_Label "Caption" "�������� ����� ������?")
		)
	)
;���������� ������ ���������
	(if (= ReloadFlag ReloadData)
		(progn
;��������� ����
			(dcl_Control_SetProperty k410odcl_ReloadDialog "TitleBarText" (strcat TitleTXT "�������� ������ ���������"))
;����������� ������ ��������� � ������
			(dcl_Control_SetProperty k410odcl_ReloadDialog_ProgressBar "Value" 0)
;����������� ������ �������
			(dcl_Control_SetProperty k410odcl_ReloadDialog_Label "Caption" "�������� ������ ���������?")
		)
	)
)

;------------------------------------------------------ ������������� ������� --

;---------------------------------------------------------------- ������ "��" --
(defun c:k410odcl_ReloadDialog_YesButton_OnClicked (/	_MSpace
														_MSItemCount
														_ReloadItem
														_CurRef
														_CurAffectiveName
														_Str
														_KeyPosition
														_ReloadCount
														_ErrorCount
														_LockCount
														_Layers)
;�����/��������� ������
	(dcl_Control_SetProperty k410odcl_ReloadDialog_YesButton "Enabled" F)
	(dcl_Control_SetProperty k410odcl_ReloadDialog_NoButton "Enabled" F)
;����������� ������ ���������
	(dcl_Control_SetProperty k410odcl_ReloadDialog_ProgressBar "Visible" T)
;�������� �����
	(dcl_Control_SetProperty k410odcl_ReloadDialog_Label "Visible" F)
	(dcl_Control_Redraw k410odcl_ReloadDialog)
;���������� ������ ������
	(if (= ReloadFlag ReloadRecordSet)
		(progn
;����� ������ ��������� (�� ���������� ������)
			(dcl_Control_SetProperty k410odcl_ReloadDialog_ProgressBar "MaxValue" 2)
;�������� �� ������ Views �������� ������ ������	
			(setq Views (vl-remove (assoc (strcat TagUDLFile ";" TagTableName ";" TagKeyFieldName) Views) Views))
;����������� ������ ���������
			(dcl_Control_SetProperty k410odcl_ReloadDialog_ProgressBar "Value" 1)
;�������� ������ ������
			(GetRecordSet)
;����������� ������ ���������
			(dcl_Control_SetProperty k410odcl_ReloadDialog_ProgressBar "Value" 2)
		)
	)
;���������� ������ ���������
	(if (= ReloadFlag ReloadData)
		(progn
;��� ������
			(setq 	_BlockRefObjectName "AcDbBlockReference"
;���������� �������� ���������
					_CurRef CurBlockRef
;��� �������� �������� ���������
					_CurAffectiveName (vla-get-EffectiveName CurBlockRef)
;������������ ������
					_MSpace (vla-get-ModelSpace ActiveDoc)
;�������� ���������� ��������� � ������������ ������			
					_MSItemCount (vla-get-count _MSpace)
;����� ������ ������
					_Str ""
;���������� ������������ ���������
					_ReloadCount 0
;���������� ������ ��� ����������
					_ErrorCount 0
;���������� ��������� �� ������������� �����
					_LockCount 0
;���� ���������
					_Layers (vla-get-Layers ActiveDoc)
;��������� ����� ���������� ������
					ReloadDataFlag T)
	
;��������� ������ ������� ������ ���������
			(dcl_Control_SetProperty k410odcl_ReloadDialog_ProgressBar "MaxValue" _MSItemCount)
;����������� ������ � ������
			(dcl_Control_SetProperty k410odcl_ReloadDialog_ProgressBar "Value" ItemIndex)
;���������� ���������
			(while (and ReloadDataFlag (/= ItemIndex _MSItemCount))
;�������� ������ �� ��������� ������� �� ������������ ������ ��� ����������� �����
				(setq _ReloadItem (vla-item _MSpace ItemIndex))
;����������, ���� ������� -- ��������� ����� ���������� ��������
				(if (and 	(= (vla-get-ObjectName _ReloadItem) _BlockRefObjectName)
							(= (vla-get-EffectiveName _ReloadItem) _CurAffectiveName))
;�������� ���� �� ����������
					(if (= (vla-get-Lock (vla-Item _Layers (vla-get-Layer _ReloadItem))) :vlax-true)	;then
						(setq _LockCount (1+ _LockCount))	;then
						(progn	;else
;������ �������� �������� ������������ ������
							(setq 	CurBlockRef _ReloadItem
;��������� ������ ������ ��������� ���� ������
									_Str "")
;������������� ��������� �����
							(InitBlock)
;��������� �������� ��������� ����
							(if (= TagKeyFieldValue "")
								(progn	;then
;���������� ���������
									(if (= (dcl_Control_GetProperty k410odcl_ChoiceRecordDialog_FillErrorCheckBox "Value") 1)
										(progn
											(repeat (length CurColumnList) 
												(setq 	_Str 	(strcat _Str 
																	(dcl_Control_GetProperty k410odcl_ChoiceRecordDialog_FillErrorTextBox "Text")
																	"\t")))
;������� ������� �� ���������� ��������
											(dcl_Grid_Clear k410odcl_ChoiceRecordDialog_Grid)
;������ � ������� �������������� ������
											(dcl_Grid_AddString k410odcl_ChoiceRecordDialog_Grid _Str)
;������ �������� � ��������
											(WriteToAttrs)
										)
									)
;���� ������
									(setq _ErrorCount (1+ _ErrorCount))
								)
								(progn	;else
;���������� ������� ��������� ���� � ������ �������� ��������� ����
									(setq _KeyPosition (vl-position TagKeyFieldValue CurKeyValueList))
									(if (null _KeyPosition)
										(progn	;then
;���������� ���������
											(if (= (dcl_Control_GetProperty k410odcl_ChoiceRecordDialog_FillErrorCheckBox "Value") 1)
												(progn
													(repeat (length CurColumnList) 
														(setq _Str (strcat 
																	_Str 
																	(dcl_Control_GetProperty k410odcl_ChoiceRecordDialog_FillErrorTextBox "Text")
																	"\t")))
;������� ������� �� ���������� ��������
													(dcl_Grid_Clear k410odcl_ChoiceRecordDialog_Grid)
;������ � ������� �������������� ������
													(dcl_Grid_AddString k410odcl_ChoiceRecordDialog_Grid _Str)
;���������� ������� �� ������ "Ok"
													(c:k410odcl_ChoiceRecordDialog_OkButton_OnClicked)
												)
											)
;���� ������
											(setq _ErrorCount (1+ _ErrorCount))
										)
										(progn	;else
;����������� ������� �� ����������� �������
											(vlax-invoke-method CurRecordSet "Move" _KeyPosition 1)
;�������� ������
											(setq 	_Str (vlax-invoke-method CurRecordSet "GetString" 2 1 "\t" "\t" nil))
;������ � ������ ������������� ������������	"\t\t" �� "\t \t"
											(repeat (length CurColumnList) (setq _Str (vl-string-subst "\t \t" "\t\t" _Str)))
;������� ������� �� ���������� ��������
											(dcl_Grid_Clear k410odcl_ChoiceRecordDialog_Grid)
;������ � ������� ����������� ������
											(dcl_Grid_AddString k410odcl_ChoiceRecordDialog_Grid _Str)
;���������� ������� �� ������ "Ok"
											(c:k410odcl_ChoiceRecordDialog_OkButton_OnClicked)
										)
									)
								)
							)
							(setq _ReloadCount (1+ _ReloadCount))
						)
					)
				)
;������� �� ��������� �������
				(setq ItemIndex (1+ ItemIndex))
;������������ �������� ������ ���������
				(dcl_Control_SetProperty k410odcl_ReloadDialog_ProgressBar "Value" ItemIndex)
			)
;����� ����� ���������� ������
			(setq 	ReloadDataFlag F
;������������� ������ �� �������� ���������
					CurBlockRef _CurRef)
;������������� ��������� �����
			(InitBlock)
;������� ������� �� ���������� ��������
			(dcl_Grid_Clear k410odcl_ChoiceRecordDialog_Grid)
;��������� �������� ��������� ����
			(if (= TagKeyFieldValue "")
;then
				(dcl_Control_SetProperty k410odcl_ChoiceRecordDialog_KeyFieldTextBox "Text" KeyFieldValueTXT)
;else
				(progn
					(dcl_Control_SetProperty k410odcl_ChoiceRecordDialog_KeyFieldTextBox "Text" TagKeyFieldValue)
;���������� ������� ��������� ���� � ������ �������� ��������� ����
					(setq _KeyPosition (vl-position TagKeyFieldValue CurKeyValueList))
					(if (null _KeyPosition)
;then
;������
						(princ)
;else
						(progn
;����������� ������� �� ����������� �������
							(vlax-invoke-method CurRecordSet "Move" _KeyPosition 1)
;�������� ������
							(setq 	_Str (vlax-invoke-method CurRecordSet "GetString" 2 1 "\t" "\t" nil))
;������ � ������ ������������� ������������	"\t\t" �� "\t \t"
							(repeat (length CurColumnList) (setq _Str (vl-string-subst "\t \t" "\t\t" _Str)))
;������� ������� �� ���������� ��������
							(dcl_Grid_Clear k410odcl_ChoiceRecordDialog_Grid)
;������ � ������� ����������� ������
							(dcl_Grid_AddString k410odcl_ChoiceRecordDialog_Grid _Str)
						)
					)
				)
			)
;��������� ������
			(alert 	(strcat "��������:		" (itoa ItemIndex) "\n"
							"���������:		" (itoa (+ _ReloadCount _LockCount)) "\n"
							"������:			" (itoa _ErrorCount) "\n"
							"�� ������������� �����:	" (itoa _LockCount)))
;����� ��������� ���������, ���� ����������� ��� �������� ���������
			(if (= ItemIndex _MSItemCount)
				(setq ItemIndex 0))
		)
	)
;�����/��������� ������
	(dcl_Control_SetProperty k410odcl_ReloadDialog_YesButton "Enabled" T)
	(dcl_Control_SetProperty k410odcl_ReloadDialog_NoButton "Enabled" T)
;�������� ������ ���������
	(dcl_Control_SetProperty k410odcl_ReloadDialog_ProgressBar "Visible" F)
;����������� �����
	(dcl_Control_SetProperty k410odcl_ReloadDialog_Label "Visible" T)
;������� ������
	(dcl_Form_Close k410odcl_ReloadDialog)
)
;---------------------------------------------------------------- ������ "��" --

;--------------------------------------------------------------- ������ "���" --
(defun c:k410odcl_ReloadDialog_NoButton_OnClicked (/)
;������� ������
	(dcl_Form_Close k410odcl_ReloadDialog)
)
;--------------------------------------------------------------- ������ "���" --
;######################################################### ������: ���������� ##

(princ)
;http://www.askit.ru/custom/progr_admin/m13/13_01_ado_basics.htm