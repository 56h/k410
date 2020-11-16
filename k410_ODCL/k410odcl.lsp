;-------------------------------------------------------------------------------
; k410 ODCL
;-------------------------------------------------------------------------------

;-------------------------------------------------------------------------------
;Расширенная информация хранится в атрибутах описания блока.
;Описание блока содержит следующие атрибуты:
; * K410_ODCL_UDLFILE - полное имя (включая путь) файла UDL источника данных;
; * K410_ODCL_TABLENAME - имя таблицы источника данных;
; * K410_ODCL_KEYFIELDNAME - имя ключевого поля;
; * K410_ODCL_ATTRNAME_1 ... K410_ODCL_ATTRNAME_N - соответствие атрибуту блока
;												поля таблицы источника данных
;Атрибуты невидимые (Invisible = T), постоянные (Constant = T) и с позицией
;фиксации (LockPosition = T).
;
; * K410_ODCL_KEYFIELDVALUE - 	значение ключевого поля таблицы источника данных
;								для выбранного вхождения.
;Атрибут невидимый (Invisible = T) и с позицией фиксации (LockPosition = T).
;-------------------------------------------------------------------------------
(setq k410_Ver "\nk410-ODCL Ver.0.06.12.07.20\n")
(princ k410_Ver)
;Загрузка расширения AciveX ACAD-a
(vl-load-com)
;Загрузка расширения ADOLisp
(load "ADOLISP_Library.lsp")
;Загрузка проекта
(command "OPENDCL")
(dcl_Project_Load "k410odcl" T)

;------------------------------------------------------ Глобальные переменные --
(setq
; - указатель приложения
		ACADApp (vlax-get-acad-object)
; - указатель на активный документ
		ActiveDoc (vla-get-ActiveDocument ACADApp)
; - полное имя файла источника данных:
		TagUDLFile ""
; - имя таблицы источника данных:
		TagTableName ""
; - ключевое поле таблицы источника данных:
		TagKeyFieldName ""
; - значение ключевого поля таблицы источника данных:
		TagKeyFieldValue ""
; - атрибут значения ключевого поля таблицы источника данных:
		AttrRefKeyFieldValue nil
; - указатель на текущее описание блока:
		CurBlock nil
; - указатель на текущее вхождение блока:
		CurBlockRef nil
; - список ссылок на постоянные атрибуты блока:
		ConstAttrRefList nil
; - список постоянных атрибутов блока:
		ConstAttrList nil
; - спсиок ссылок на видимые атрибуты блока:
		AttrRefList nil
; - список видимых атибутов блока:
		AttrList nil
; - список полей таблицы источника данных, соответствующих атрибутам:
		AttrFieldsList nil
; - наборы данных:
		Views nil
; - список полей таблицы текщуего источника данных:
		CurColumnList nil
; - текущий список значений ключевого поля:
		CurKeyValueList nil
; - текущий набор данных:
		CurRecordSet nil
; - флаг обновления данных:
		ReloadDataFlag F
; - флаг создания новых атрибутов:		
		NewAttrFlag F
; - индекс элемента обновления:
		ItemIndex 0
; - флаг диалога обновления:
		ReloadFlag 0
; - список значений ключей:
		KeyValues nil
; - строка значений ключей:
		KeyValuesStr ""
; - разделитель между значениями ключей в строке значений ключей:
		StrDelimiter ", "
; - флаг сброса списка и строки значений:
		KeyValuesClear T
;------------------------------------------------------ Глобальные переменные --

;------------------------------------------------------- Переменные-константы --
		AttrUDLFile "K410_ODCL_UDLFILE"
		AttrTableName "K410_ODCL_TABLENAME"
		AttrKeyFieldName "K410_ODCL_KEYFIELDNAME"
		AttrKeyFieldValue "K410_ODCL_KEYFIELDVALUE"
		AttrPrefix "K410_ODCL_"

		FieldTXT "Выберите поле..."
		TableViewTXT "Выберите таблицу или запрос..."
		KeyFieldTXT "Выберите ключевое поле..."
		KeyFieldValueTXT "Значение ключевого поля..."
		ErrorTitleTxt "k410 ODCL - Ошибка"
		TitleTXT "k410 ODCL - "
		
		ArrayWidth "10"
		ArraydX "100"
		ArraydY "50"
		
;Обновить источник данных
		ReloadRecordSet 1
;Обновить данные атрибутов
		ReloadData 2
;------------------------------------------------------- Переменные-константы --

;------------------------------------------------------------ Перечень ошибок --
		Error_001_TXT "Ошибка 001: Ошибка ключевого поля"
;					ключевое поле, сохраненное в описании блока	отсутствует в 
;					таблице источника данных
		Error_002_TXT "Ошибка 002: Ошибка открытия UDL файла источника данных"
;					UDL файла не найден по сохраненному пути
		Error_003_TXT "Ошибка 003: Ошибка загрузки наборов данных"
;					в загруженных наборах nil вместо указателя на набор, или
;					вместо списка полей набора
		Error_004_TXT "Ошибка 004: Ошибка значения ключевого поля"
;					внесенное в текстовое поле значение ключевого поля 
;					отсутствует в списке значений ключевого поля
;		Error_005_TXT
;------------------------------------------------------------ Перечень ошибок --
)


;--------------------------------------- Команда отображения версии программы --
(defun c:k410_Ver (/)
	k410_Ver
)	
;--------------------------------------- Команда отображения версии программы --

;---------------------------------------------------- Добавление пунктов меню --

;--------------------------- Функция поиска элемента меню по имени (по Label) --
;Вход:
; * _Menu - меню, в котором ищется пункт;
; * _LabelString - имя пункта меню;
;Выход:
; * номер элемента, если такая метка есть;
; * nil, если такого элемента нет;
(defun FindSimpleLabel (_Menu
						_LabelString /	_MenuItemCount
										_i
										_Result)
;Количество элементов меню
	(setq	_MenuItemCount (vla-Get-Count _Menu)
;Обнуление индекса элементов меню
			_i 0
;Результат работы функции
			_Result nil)
	(repeat _MenuItemCount
		(if (= (vla-Get-Label (vla-Item _Menu _i)) _LabelString)
			(setq _Result _i))
		(setq _i (1+ _i))
	)
;Возврат значения
	_Result
)
;--------------------------- Функция поиска элемента меню по имени (по Label) --

;-------------------------------------------- Функция добавления пунктов меню --
(defun AddMenuItems (	_Menu /	_Item_1
                              	_Item_2
								_Item_3)
	(vla-AddSeparator _Menu 0)

	(setq	_Item_3 (vla-AddMenuItem _Menu 0 "Массив вхождений" "(BlockRefArray) ")
			_Item_2 (vla-AddMenuItem _Menu 0 "Выбор записи" "(ChoiceRecord) ")
			_Item_1 (vla-AddMenuItem _Menu 0 "Адреса данных" "(DataAddr) "))

	(if _Item_1 (vla-Put-HelpString _Item_1 "Создание или редактирование адресов данных"))
	(if _Item_2 (vla-Put-HelpString _Item_2 "Выбор записи таблицы источника данных"))
	(if _Item_3 (vla-Put-HelpString _Item_3 "Создание массива вхождений по всем записям источника данных"))
	
	(princ)
)
;-------------------------------------------- Функция добавления пунктов меню --
;Загрузка указателя на PopupMenu
(setq	_MenuGroup (vla-Item (vla-Get-MenuGroups ACADApp) "ACAD")
		_PopupMenu (vla-Item (vla-Get-Menus _MenuGroup) "Меню объектов блоков с атрибутами"))

;Добавление пунктов меню, если они не были добавлены
(if (FindSimpleLabel _PopupMenu "Адреса данных")
;then
	(princ)
;else
	(AddMenuItems _PopupMenu)
)
;---------------------------------------------------- Добавление пунктов меню --

;----------------------------------------------------------------- ODCL часть --

;---------------------------- Функция отображения сообщения об ошибке ADOLISP --
;Функци является переработкой функции ADOLISP_ErrorPrinter пакета
;ADOLISP_Library.lsp

(defun k410_ErrorPrinter (/ _ErrorString)
;Инициализация переменных
	(setq _ErrorString "")
;Формирование строки с последним запросом
	(if ADOLISP_LastSQLStatement
		(setq _ErrorString (strcat	"Last SQL statement:\n\""
									ADOLISP_LastSQLStatement
									"\"\n\n")
		)
	)
;Извлечение информации об ошибке
	(foreach ErrorList ADOLISP_ErrorList
		(foreach ErrorItem ErrorList
			(setq _ErrorString
				(strcat _ErrorString (car ErrorItem) ": " (cdr ErrorItem) "\n")
			)
		)
	)
;Отображение сообщения в выпадающем окне
	(dcl_MessageBox _ErrorString ErrorTitleTxt 2 4)
;"Тихий" выход
	(princ)
)
;---------------------------- Функция отображения сообщения об ошибке ADOLISP --

;------------------------------- Функция загрузки данных из источника в набор --
;Является переработкой функции ADOLISP_DoSQL пакета ADOLISP_Library.lsp
;Поскольку выполняем только запрос SELECT ... объект Command не применяем
;Вход:
; - ConnectionObject - соединение,
; - SQLStatement - текст запроса.
;Выход:
; - nil - в случае путсого или ошибочного набора данных;
; - список - ((список полей набора данных) vla-object НаборДанных)

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
;Сброс переменных
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
;------------------------------- Функция загрузки данных из источника в набор --

;------------------------------------------- Функция поиска элемента в списке --

;----------------------------- Функция загрузки атрибутов из описания блока c --
;							   заданным логическим свойством 
;Вход:
; *	_BlockRef - ссылка на вхождение блока;
; *	_AttrProperty - строка с именем логического свойства:
;	- Backward;
;	- Constant;
;	- HasExtensionDictionary;
;	- Invisible;
;	- LockPosition;
;	- MTextAttribute;
;	- Preset;
;	- UpsideDown;
;	- Verify;
; *	_EqFlag - значение переменной (:vlax-true или :vlax-false);
;Выход:
; *	список ссылок на атрибуты из описания блока
; *	nil в случае отсутствия атрибутов у блока
(setq 	AttrObjectName 		"AcDbAttributeDefinition")

(defun GetAttributes (	_BlockRef 
						_AttrProperty 
						_EqFlag /	_TempVar
									_Block
									_Index
									_IntermediateResult
									_Result)
;Сброс и установка переменных
	(setq 	_Index 0
			_Result nil
			_Block (vla-Item (vla-get-Blocks (vla-get-ActiveDocument (vlax-get-acad-object))) (vla-get-EffectiveName _BlockRef)))
;Поиск объектов "атрибут" со свойством "видимый"
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
;							   Функция загрузки атрибутов из описания блока c
;----------------------------- заданным логическим свойством -------------------

;------------------ Функция инициализации переменных для выбранного вхождения --
(defun InitBlock (/	_Index
					_TempVar)
;Получение указателя на выбронное вхождение
	(if (not ReloadDataFlag)
		(setq CurBlockRef (vla-Item (vla-get-PickfirstSelectionSet ActiveDoc) 0))
	)
;Получение указателя на описание блока
	(setq 	CurBlock (vla-Item (vla-get-Blocks ActiveDoc) (vla-get-EffectiveName CurBlockRef))
;Загрузка списка ссылок на рабочие атрибуты блока
			AttrRefList (GetAttributes CurBlockRef "Constant" :vlax-false)
;Загрузка списка ссылок на константные атрибуты блока
			ConstAttrRefList (GetAttributes CurBlockRef "Constant" :vlax-true)
;Сброс переменных
			ConstAttrList nil
			AttrList nil
			AttrFieldsList nil
			AttrRefKeyFieldValue nil
			TagUDLFile ""
			TagTableName ""
			TagKeyFieldName ""
			TagKeyFieldValue ""
			NewAttrFlag F)
;Формирование списка имен константных атрибутов
	(foreach _ConstAttr ConstAttrRefList
		(setq ConstAttrList (append ConstAttrList (list (vla-get-TagString _ConstAttr))))
;Запись значения UDLFile
		(if (= (vla-get-TagString _ConstAttr) AttrUDLFile)
			(setq TagUDLFile (vla-get-TextString _ConstAttr));then
		)
;Запись значения имени таблицы источника данных	
		(if (= (vla-get-TagString _ConstAttr) AttrTableName)
			(setq TagTableName (vla-get-TextString _ConstAttr));then
		)
;Запись имени ключевого поля таблицы источника данных		
		(if (= (vla-get-TagString _ConstAttr) AttrKeyFieldName)
			(setq TagKeyFieldName (vla-get-TextString _ConstAttr));then
		)		
	)
;Формирование списка имен рабочих атрибутов
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
;Поиск атрибута значения ключевого поля таблицы источника данных
	(foreach _AttrRef (vlax-safearray->list (vlax-variant-value (vla-GetAttributes CurBlockRef)))
;Загрузка значения ключевого поля
		(if (= (vla-Get-TagString _AttrRef) AttrKeyFieldValue)
;then
			(setq 	TagKeyFieldValue (vla-Get-TextString _AttrRef)
					AttrRefKeyFieldValue _AttrRef)
		)
	)
)
;------------------ Функция инициализации переменных для выбранного вхождения --

;----------------------------------------- Функция формирования набора данных --
;Выход:
; - CurColumnList - список имен полей загруженного набора данных
; - CurRecordSet - ссылка на загруженный набор данных

(defun GetRecordSet (/	_KeyFields
						_KeyValues
						_FullView
						_ConnectionObj)

;Поиск сохраненного набора данных
	(setq 	_View (assoc (strcat TagUDLFile ";" TagTableName ";" TagKeyFieldName) Views)
;Сброс переменных
			CurRecordSet nil
			CurColumnList nil
			CurKeyValueList nil
			;ADOLISP_DoNotForceJetODBCParsing nil
			)
	(if (null _View)
;then
;Подключение источника данных
		(if (/= TagUDLFile "")
;then
			(if (findfile TagUDLFile)
;Подключение к источнику данных
				(progn
					(prompt "Открытие соединения с источником данных\n")
					(setq _ConnectionObj (ADOLISP_ConnectToDB TagUDLFile "" ""))
					(if (null _ConnectionObj)
;then
						(k410_ErrorPrinter)
;else
						(progn
							(prompt "Загрузка набора данных\n")
							(setq	_KeyFields (last (k410_DoSQL _ConnectionObj (strcat "SELECT " TagKeyFieldName " FROM " TagTableName ";")))
									_FullView (k410_DoSQL _ConnectionObj (strcat "SELECT * FROM " TagTableName ";")))
							(prompt "Загрузка значений ключевого поля\n")
;Позиционирование курсора на первую запись набора данных
							(vlax-invoke-method _KeyFields "MoveFirst")
;Загрузка значений ключевого поля
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
;Закрытие набора данных значений ключевого поля
							(vlax-invoke-method _KeyFields "Close")
							(vlax-release-object _KeyFields)
;Формирование итогового списка							
							(setq	_View (list (car _FullView) 
												_KeyValues
												(last _FullView)))													
;Проверка на отсутствие nil-ов в полученных списках
							(if (null (vl-position nil _View))
;then								
								(progn
;Список полей таблицы источника данных
									(setq 	CurColumnList (car _View)
;Список значений ключевого поля
											CurKeyValueList _KeyValues
;Набор данных
											CurRecordSet (last _FullView)
;Добавление набора данных к списку наборов
											Views (append 	Views 
															(list (append (list (strcat TagUDLFile ";" TagTableName ";" TagKeyFieldName)) _View))))
;Проверка ключевого поля
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
;Список полей таблицы источника данных
		(setq 	CurColumnList (cadr _View)
;Список значений ключевого поля
				CurKeyValueList (caddr _View)
;Набор данных
				CurRecordSet (last _View))
	)
)
;----------------------------------------- Функция формирования набора данных --

;------------------------------------------- Функция обновления набора данных --
(defun ReloadRecordSet (/)
;Установка флага диалога обновления
	(setq ReloadFlag ReloadRecordSet)
;Диалог обновления набора данных
	(prompt "\n")
;Отображение формы
	(dcl_Form_Show k410odcl_ReloadDialog)
;Сброс флага диалога обновления
	(setq ReloadFlag 0)
)
;------------------------------------------- Функция обновления набора данных --

;------------------------------------------ Команды извлечения значений ключа --
;Выход:
; - nil - атрибута ключа нет или значение ключа равно nil
; - список значений KeyValues и в строка значений KeyValuesStr (полученную 
;	строку	удобно использовать при посроении SQL-запросов к источникам данных)
;
;Если установлен флаг KeyValuesClear (= Т) перед выполнением команды списки
;значений сбрасываются

;Команда получения массива значений ключей из текущего выделения
(defun GetKeyValues (/	_Index
							_PFSS
							_Item)
;Инициализация переменных
	(setq 	_Index 0
			ReloadDataFlag T
			_PFSS (vla-get-PickFirstSelectionSet ActiveDoc))
;Сброс списков значений ключей
	(if (= KeyValuesClear T)
		(setq 	KeyValues nil
				KeyValuesStr "")
	)
;Извлечение значений ключевого поля
	(repeat  (vla-get-Count _PFSS)
;Загрузка ссылки на элемент выделения
		(setq _Item (vla-Item _PFSS _Index))
;Получение значения ключевого поля
		(if (= (vla-get-ObjectName _Item) "AcDbBlockReference")
			(if (= (vla-get-HasAttributes _Item) :vlax-true)
				(progn
					(setq CurBlockRef _Item)
;Инициализация выбранного блока
					(InitBlock)
					(if (/= TagKeyFieldValue "")
						(progn
;Добавление значения ключа в список
							(setq KeyValues (append KeyValues (list TagKeyFieldValue)))
;Добавление значения ключа в строку значений
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
;Сброс флага обновления данных
	(setq ReloadDataFlag F)
)
;Длинная команда
(defun c:k410_GetKeyValues (/)
	(GetKeyValues)
)
;Короткая команда
;Перед выполнением команды список и строка значений сбрасываются
(defun c:k410_gkv (/)
	(GetKeyValues)
;Возврат загруженного списка
	KeyValuesStr
)
;-------------------------------------------------------------------------------
;Сброс списка значений ключей KeyValues и KeyValuesStr
(defun c:k410_KeyValuesClear (/)
;Сброс списка и строки значений
	(setq 	KeyValues nil
			KeyValuesStr "")
)
;------------------------------------------ Команды извлечения значений ключа --

;----------------------------------------------------- Функция OR для списков --
;Вход:
;	- _L1, _L2 - списки
;Выход:
;	список, состоящий из элементов _L1 и элементов _L2
(defun ListOR (_L1 _L2 /)
	(if (= (and (listp _L1) (listp _L2)) T)
;Объединение списков и возвращение результата
		(append _L1 _L2);then
		(prompt "\nАргументами функции должны быть списки\n");else
	)
)
;----------------------------------------------------- Функция OR для списков --

;---------------------------------------------------- Функция XOR для списков --
;Вход:
;	- _L1, _L2 - списки
;Выход:
;	список, состоящий из элементов _L1, которых нет в _L2 и элементов _L2,
;	которых нет в _L1
(defun ListXOR (	_L1 _L2 /	_Item
								_FullList
								_Part1
								_Part2)
	(if (= (and (listp _L1) (listp _L2)) T)
		(progn
;Cписок, состоящий из всех элементов списков
			(setq 	_FullList (append _L1 _L2)
					_Part1 _FullList
					_Part2 _FullList)
;Удаление элементов первого списка
			(foreach _Item _L1
				(setq _Part1 (vl-remove _Item _Part1))
			)
;Удаление элементов второго списка
			(foreach _Item _L2
				(setq _Part2 (vl-remove _Item _Part2))
			)
;Объединение полученных частей и возвращение результата
			(append _Part2 _Part1)
		)
		(prompt "\nАргументами функции должны быть списки\n")
	)
)
;---------------------------------------------------- Функция XOR для списков --

;---------------------------------------------------- Функция AND для списков --
;Вход:
;	- _L1, _L2 - списки
;Выход:
;	список, состоящий из элементов, которые есть и в_L1 и в _L2
(defun ListAND (	_L1 _L2 /	_Item
								_FullList
								_Part)
	(if (= (and (listp _L1) (listp _L2)) T)
		(progn
;Cписок, состоящий из всех элементов списков
			(setq 	_FullList (append _L1 _L2)
					_Part _FullList)
;Удаление элементов второго списка 
;(получение списка элементов _L1, которых нет в _L2)
			(foreach _Item _L2
				(setq _FullList (vl-remove _Item _FullList))
			)
;Удаление элементов, которых нет во втором списке
			(foreach _Item _FullList
				(setq _L1 (vl-remove _Item _L1))
			)
;Возвращение результата
			_L1
		)
		(prompt "\nАргументами функции должны быть списки\n")
	)
)
;---------------------------------------------------- Функция AND для списков --

;###################################################### Диалог: адреса данных ##
;------------------------- Функция создания или редактирования адресов данных --
(defun DataAddr (/)
	(prompt "\n")
	(if (= (vla-get-ActiveSpace ActiveDoc) 1)
		(dcl_Form_Show k410odcl_DataAddrDialog);then
		(alert "Функция реализована только для объектов пространства модели");else
	)
	(princ)
;Обновить, если созданы новые атрибуты
	(if NewAttrFlag 
		(progn
			(command "_.battman"))
			(setq NewAttrFlag F)
		)
)
;------------------------- Функция создания или редактирования адресов данных --

;------------------------------------------------------ Инициализация диалога --
(defun c:k410odcl_DataAddrDialog_OnInitialize (/ _Index)
;Инициализация переменных
	(InitBlock)
;Сброс элементов и флагов управления
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
;Сброс переменных
	(setq _Index 0)
;Запись списка имен видимых атрибутов
	(repeat (length AttrList)
		(dcl_Grid_AddRow k410odcl_DataAddrDialog_Grid (nth _Index AttrList) (nth _Index AttrFieldsList))
		(if (/= (nth _Index AttrFieldsList) "") (dcl_Control_SetProperty k410odcl_DataAddrDialog_DelRecordButton "Enabled" T))
		(setq _Index (1+ _Index))
	)
;Загрузка набора данных
	(GetRecordSet)
	(if (not (null CurColumnList))
		(progn
			(dcl_Control_SetProperty k410odcl_DataAddrDialog_FieldNameLabel "Enabled" T)
			(dcl_Control_SetProperty k410odcl_DataAddrDialog_FieldComboBox "Enabled" T)
			(dcl_Control_SetProperty k410odcl_DataAddrDialog_FieldComboBox "List" CurColumnList)
		)
	) 
)
;------------------------------------------------------ Инициализация диалога --

;----------------------------------------------------- Событие изменения поля --
(defun c:k410odcl_DataAddrDialog_FieldComboBox_OnSelChanged (ItemIndexOrCount Value /)
;Включение кнопки "Добавить запись"
	(dcl_Control_SetProperty k410odcl_DataAddrDialog_AddRecordButton "Enabled" T)
)
;----------------------------------------------------- Событие изменения поля --

;--------------------------------------------------- Кнопка "Источник данных" --
(defun c:k410odcl_DataAddrDialog_DSNameButton_OnClicked (/ _Index)
;Открытие диалога по выбору нового источника данных
	(dcl_Form_Show k410odcl_OpenUDLForm)
;Сброс элементов управления и очистка адресов данных, если изменился источник
;данных
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
;--------------------------------------------------- Кнопка "Источник данных" --

;--------------------------------------------------- Кнопка "Добавить запись" --
(defun c:k410odcl_DataAddrDialog_AddRecordButton_OnClicked (/	_Attr
																_Field
																_Row)
;Загрузка имени поля
	(setq 	_Field (dcl_Control_GetProperty k410odcl_DataAddrDialog_FieldComboBox "Text")
;Получение номера строки
			_Row (car (dcl_Grid_GetCurCell k410odcl_DataAddrDialog_Grid)))
;Сохранение записи в таблице
	(if (/= _Row -1)
;then
		(progn 
			(dcl_Grid_SetCellText k410odcl_DataAddrDialog_Grid _Row 1 _Field)
;Отображение кнопки "Удалить запись"
   
			(dcl_Control_SetProperty k410odcl_DataAddrDialog_DelRecordButton "Enabled" T)
		)
	)
;Отображение кнопки "Сохранить"
	(dcl_Control_SetProperty k410odcl_DataAddrDialog_SaveButton "Enabled" T)
)
;--------------------------------------------------- Кнопка "Добавить запись" --

;---------------------------------------------------- Кнопка "Удалить запись" --
(defun c:k410odcl_DataAddrDialog_DelRecordButton_OnClicked (/	_Row
																_СellList)
;Получение номера строки
	(setq _Row (car (dcl_Grid_GetCurCell k410odcl_DataAddrDialog_Grid)))
;Очистка записи
	(dcl_Grid_SetCellText k410odcl_DataAddrDialog_Grid _Row 1 "")
;Включение кнопки "Сохранить"
	(dcl_Control_SetProperty k410odcl_DataAddrDialog_SaveButton "Enabled" T)
;Проверка на пустые записи и отключение кнопки, если все записи удалены
	(setq _СellList (dcl_Grid_GetColumnCells k410odcl_DataAddrDialog_Grid 1))
	(dcl_Control_SetProperty k410odcl_DataAddrDialog_DelRecordButton "Enabled" F)
	(repeat (length _СellList)
		(if (/= (car _СellList) "")
;then
			(dcl_Control_SetProperty k410odcl_DataAddrDialog_DelRecordButton "Enabled" T)
		)
		(setq _СellList (cdr _СellList))
	)
;Отображение кнопки "Сохранить"
	(dcl_Control_SetProperty k410odcl_DataAddrDialog_SaveButton "Enabled" T)  
)
;---------------------------------------------------- Кнопка "Удалить запись" --

;--------------------------------------------------------- Кнопка "Сохранить" --
(defun c:k410odcl_DataAddrDialog_SaveButton_OnClicked (/	_AttrIndex
															_Index
															_AttrList
															_FieldList
															_AttrName)
;Поиск атрибута UDLFile
	(setq _AttrIndex (vl-position AttrUDLFile ConstAttrList))
;Запись значения UDLFile
	(if (null _AttrIndex)
;then
;Создание атрибута UDL файла
		(setq 	_AttrIndex 
				(vla-AddAttribute	CurBlock
									0.001																				;Height
									(+ 	acAttributeModeInvisible														;Mode
									acAttributeModeConstant
									acAttributeModeLockPosition)
									"k410 ODCL: служебный атрибут"														;Prompt
									(vlax-safearray-fill (vlax-make-safearray vlax-vbDouble '(0 . 2)) '(0.0 0.0 0.0))	;Insertion point
									AttrUdlFile																			;Tag
									TagUDLFile																			;Value
				)
				ConstAttrList (append ConstAttrList (list AttrUdlFile))
				ConstAttrRefList (append ConstAttrRefList (list _AttrIndex))
;Установка флага создания нового атрибута
				NewAttrFlag T
		)
;else
;Запись в атрибут UDL файла
		(vla-put-textstring (nth _AttrIndex ConstAttrRefList) TagUDLFile)
	)
;Поиск атрибута таблицы источника данных
	(setq _AttrIndex (vl-position AttrTableName ConstAttrList))
;Запись значения имени таблицы источника данных
	(if (null _AttrIndex)
;then
;Создание атрибута имени таблицы источника данных
		(setq 	_AttrIndex 
				(vla-AddAttribute	CurBlock
									0.001																				;Height
									(+ 	acAttributeModeInvisible														;Mode
										acAttributeModeConstant
										acAttributeModeLockPosition)
									"k410 ODCL: служебный атрибут"														;Prompt
									(vlax-safearray-fill (vlax-make-safearray vlax-vbDouble '(0 . 2)) '(0.0 0.0 0.0))	;Insertion point
									AttrTableName																		;Tag
									TagTableName																		;Value
				)	 
				ConstAttrList (append ConstAttrList (list AttrTableName))
				ConstAttrRefList (append ConstAttrRefList (list _AttrIndex))
;Установка флага создания нового атрибута
				NewAttrFlag T
		)
;else
;Запись в атрибут имени таблицы источника данных
		(vla-put-textstring (nth _AttrIndex ConstAttrRefList) TagTableName)
	)
;Поиск атрибута ключевого поля таблицы источника данных
	(setq _AttrIndex (vl-position AttrKeyFieldName ConstAttrList))
;Запись имени ключевого поля таблицы источника данных
	(if (null _AttrIndex)
;then
;Создание атрибута ключевого поля таблицы источника данных
		(setq 	_AttrIndex 
				(vla-AddAttribute	CurBlock
									0.001																				;Height
									(+ 	acAttributeModeInvisible														;Mode
										acAttributeModeConstant
										acAttributeModeLockPosition)
									"k410 ODCL: служебный атрибут"														;Prompt
									(vlax-safearray-fill (vlax-make-safearray vlax-vbDouble '(0 . 2)) '(0.0 0.0 0.0))	;Insertion point
									AttrKeyFieldName																	;Tag
									TagKeyFieldName																		;Value
				)
				ConstAttrList (append ConstAttrList (list AttrKeyFieldName))
				ConstAttrRefList (append ConstAttrRefList (list _AttrIndex))
;Установка флага создания нового атрибута
				NewAttrFlag T
		)
;else
;Запись в атрибут ключевого поля таблицы источника данных
		(vla-put-textstring (nth _AttrIndex ConstAttrRefList) TagKeyFieldName)
	)
;Запись значения ключевого поля таблицы источника данных
	(if (null AttrRefKeyFieldValue)
;then
;Создание атрибута ключевого поля таблицы источника данных
		(setq 	_AttrIndex 
				(vla-AddAttribute	CurBlock
									0.001																				;Height
									(+ 	acAttributeModeInvisible														;Mode
										acAttributeModeLockPosition)
									"k410 ODCL: служебный атрибут"														;Prompt
									(vlax-safearray-fill (vlax-make-safearray vlax-vbDouble '(0 . 2)) '(0.0 0.0 0.0))	;Insertion point
									AttrKeyFieldValue																	;Tag
									TagKeyFieldValue																	;Value
				)			
				AttrRefKeyFieldValue _AttrIndex
;Установка флага создания нового атрибута
				NewAttrFlag T
		)
;else
;Запись в атрибут ключевого поля таблицы источника данных
		(vla-put-textstring AttrRefKeyFieldValue TagKeyFieldValue)
	)
;Инициализацтя пременных
	(setq 	_Index 0
			_AttrList (dcl_Grid_GetColumnCells k410odcl_DataAddrDialog_Grid 0)
			_FieldList (dcl_Grid_GetColumnCells k410odcl_DataAddrDialog_Grid 1)
;Запись адресов данных
			_Index 0)
	(repeat (length _AttrList)
		(setq 	_AttrName (strcat AttrPrefix (nth _Index _AttrList))
				_AttrIndex (vl-position _AttrName ConstAttrList))
		(if (null _AttrIndex)
;then
;Создание атрибута ключевого поля таблицы источника данных
			(setq 	_AttrIndex 
					(vla-AddAttribute	CurBlock
										0.001																				;Height
										(+ 	acAttributeModeInvisible														;Mode
											acAttributeModeConstant
											acAttributeModeLockPosition)
										"k410 ODCL: служебный атрибут"														;Prompt
										(vlax-safearray-fill (vlax-make-safearray vlax-vbDouble '(0 . 2)) '(0.0 0.0 0.0))	;Insertion point
										_AttrName																			;Tag
										(nth _Index _FieldList)																;Value
					)
					ConstAttrList (append ConstAttrList (list _AttrName))
					ConstAttrRefList (append ConstAttrRefList (list _AttrIndex))
;Установка флага создания нового атрибута
					NewAttrFlag T
			)
;else
;Запись в атрибут значения поля таблицы источника данных
			(vla-put-textstring (nth _AttrIndex ConstAttrRefList) (nth _Index _FieldList))
		)
		(setq _Index (1+ _Index))
	)
;Отключение кнопки "Сохранить
	(dcl_Control_SetProperty k410odcl_DataAddrDialog_SaveButton "Enabled" F)  
)
;--------------------------------------------------------- Кнопка "Сохранить" --

;----------------------------------------------------------- Кнопка "Закрыть" --
(defun c:k410odcl_DataAddrDialog_CancelButton_OnClicked (/)
	(dcl_Form_Close k410odcl_DataAddrDialog 2)
)
;----------------------------------------------------------- Кнопка "Закрыть" --

;------------------ Закрытие диалога "Открыть файл описания источника данных" --
(defun c:k410odcl_OpenUDLForm_OnClose (	UpperLeftX
										UpperLeftY /	_ConnectionObj
														_View)
;Загрузка полного имени (включая путь) файла UDL источника данных
	(setq TagUDLFile (dcl_FileExplorer_GetPathName k410odcl_OpenUDLForm_FileExplorer))
;Проверка длины имени UDL файла источника данных
	(if (> (strlen TagUDLFile) 255)
		(progn
			(dcl_MessageBox	"Длина полного имени файла UDL источника данных (включая путь)\n не должна превышать 255 символов."
							ErrorTitleTxt
							2 4)
			(setq TagUDLFile "")
		)  
	)
;Сброс флага выбора нового источника данных
	(setq _DSChangeFlag F)
;Подключение источника данных
	(if (/= TagUDLFile "")
;then
;Подключение к источнику данных
		(progn
			(prompt "Подключение к источнику данных\n")
			(setq _ConnectionObj (ADOLISP_ConnectToDB TagUDLFile "" ""))
			(if (null _ConnectionObj)
;then
				(k410_ErrorPrinter)
;else
				(progn
;Отображение диалога выбора источника данных
					(dcl_Form_Show k410odcl_DSChoiceDialog)
;Закрытие соединения
					(ADOLISP_DisconnectFromDB _ConnectionObj)
				)  
			)	
		)
;else
		(princ)
	)
)
;------------------ Закрытие диалога "Открыть файл описания источника данных" --

;---------------------- Инициализация диалога выбора таблицы и ключевого поля --
;Диалог открывается, если соединение было создано
(defun c:k410odcl_DSChoiceDialog_OnInitialize (/	_TV
													_TablesAndViews)
;Загрузка перечня таблиц и запросов

	(setq 	_TV (ADOLISP_GetTablesAndViews _ConnectionObj)
			_TablesAndViews (append (nth 0 _TV) (nth 1 _TV)))
;Запись перечня таблиц и запросов в выпадающий список
	(dcl_Control_SetProperty k410odcl_DSChoiceDialog_DSTableComboBox "List" _TablesAndViews)
;Сброс текстовых полей выпадающих списков, кнопок
	(dcl_Control_SetProperty k410odcl_DSChoiceDialog_DSTableComboBox "Text" TableViewTXT)
	(dcl_Control_SetProperty k410odcl_DSChoiceDialog_DSKeyFieldComboBox "Text" KeyFieldTXT)
	(dcl_Control_SetProperty k410odcl_DSChoiceDialog_DSKeyFieldComboBox "Enabled" F)
	(dcl_Control_SetProperty k410odcl_DSChoiceDialog_DSKeyFieldLabel "Enabled" F)
	(dcl_Control_SetProperty k410odcl_DSChoiceDialog_OkButton "Enabled" F)
)
;---------------------- Инициализация диалога выбора таблицы и ключевого поля --

;------------------------------------------------ Событие смены имени таблицы --
(defun c:k410odcl_DSChoiceDialog_DSTableComboBox_OnSelChanged (	ItemIndexOrCount 
																Value / _CL)
;Сброс текстового поля выпадающего списка
	(dcl_Control_SetProperty k410odcl_DSChoiceDialog_DSKeyFieldComboBox "Text" KeyFieldTXT)
;Получение списка полей таблицы или запроса
	(setq _CL (ADOLISP_GetColumns _ConnectionObj Value))
	(if (null _CL)
;then
		(progn
;Сообщение об ошибке
			(k410_ErrorPrinter)
;Сброс выпадающего списка
			(dcl_Control_SetProperty k410odcl_DSChoiceDialog_DSKeyFieldComboBox "Enabled" F)
			(dcl_Control_SetProperty k410odcl_DSChoiceDialog_DSKeyFieldLabel "Enabled" F)       
		)
;else		
		(progn
;Включение выпадающего списка
			(dcl_Control_SetProperty k410odcl_DSChoiceDialog_DSKeyFieldComboBox "Enabled" T)
			(dcl_Control_SetProperty k410odcl_DSChoiceDialog_DSKeyFieldLabel "Enabled" T)
;Формирование и запись выпадающего списка
			(setq CurColumnList nil)
			(foreach _ColName _CL
				(setq 	CurColumnList (append CurColumnList (list (car _ColName))))
			)
			(dcl_Control_SetProperty k410odcl_DSChoiceDialog_DSKeyFieldComboBox "List" CurColumnList)
		)
	)
)
;------------------------------------------------ Событие смены имени таблицы --

;----------------------------------------------- Событие смены ключевого поля --
(defun c:k410odcl_DSChoiceDialog_DSKeyFieldComboBox_OnSelChanged (ItemIndexOrCount Value /)
;Отображение кнопки "Ок"
	(dcl_Control_SetProperty k410odcl_DSChoiceDialog_OkButton "Enabled" T)
)
;----------------------------------------------- Событие смены ключевого поля --

;------------- Кнопка "Ok" и закрытие диалога выбора таблицы и ключевого поля --
(defun c:k410odcl_DSChoiceDialog_OkButton_OnClicked (/	_View
														_Col
														_Str
														_Prop)
;Сохранение имени таблицы и ключевого поля
	(setq 	TagTableName (dcl_Control_GetProperty k410odcl_DSChoiceDialog_DSTableComboBox "Text")
			TagKeyFieldName (dcl_Control_GetProperty k410odcl_DSChoiceDialog_DSKeyFieldComboBox "Text"))
;Запись перечня полей в выпадающий список
	(dcl_Control_SetProperty k410odcl_DataAddrDialog_FieldComboBox "List" CurColumnList)
	(dcl_Control_SetProperty k410odcl_DataAddrDialog_FieldNameLabel "Enabled" T)
	(dcl_Control_SetProperty k410odcl_DataAddrDialog_FieldComboBox "Enabled" T)
	(dcl_Control_SetProperty k410odcl_DataAddrDialog_SaveButton "Enabled" T)
	(dcl_Control_SetProperty k410odcl_DataAddrDialog_LabelUDLFilePrompt "Caption" TagUDLFile)
	(dcl_Control_SetProperty k410odcl_DataAddrDialog_LabelTableNamePrompt "Caption" TagTableName)
	(dcl_Control_SetProperty k410odcl_DataAddrDialog_LabelKeyFieldPrompt "Caption" TagKeyFieldName)  
;Закрытие диалога с кодом 1 ([Enter])
	(dcl_Form_Close k410odcl_DSChoiceDialog 1)
;Установка флага выбора нового источника данных
	(setq _DSChangeFlag T)
;Установка флага сброса адресов данных
	(if (= (dcl_Control_GetProperty k410odcl_DSChoiceDialog_CheckBox "Value") 1)
;then
		(setq _DSClearAddr F)
;else
		(setq _DSClearAddr T)
	)
)
;------------- Кнопка "Ok" и закрытие диалога выбора таблицы и ключевого поля --
;###################################################### Диалог: адреса данных ##

;####################################################### Диалог: выбор записи ##
(defun ChoiceRecord (/)
	(prompt "\n")
	(if (= (vla-get-ActiveSpace ActiveDoc) 1)
		(dcl_Form_Show k410odcl_ChoiceRecordDialog);then
		(alert "Функция реализована только для объектов пространства модели");else
	)
	(princ)
)


;--------------------------------------------------------- Заполнение таблицы --
(defun FillChoiceRecordDialogTable (/	_Str
										_KeyPosition)
;Сброс или установка кнопки "Обновить набор данных"
	(if (null CurRecordSet)
		(dcl_Control_SetProperty k410odcl_ChoiceRecordDialog_ReloadDSButton "Enabled" F)
		(dcl_Control_SetProperty k410odcl_ChoiceRecordDialog_ReloadDSButton "Enabled" T)
	)	
;Удаление колонок предыдущего набора (если они сохранились)	
	(while (/= (dcl_Grid_GetColumnCount k410odcl_ChoiceRecordDialog_Grid) 0)
		(dcl_Grid_DeleteColumn k410odcl_ChoiceRecordDialog_Grid 0)
	)
;Создание колонок таблицы
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
;Вычисление позиции ключевого поля в списке значений ключевого поля
			(setq _KeyPosition (vl-position TagKeyFieldValue CurKeyValueList))
			(if (null _KeyPosition)
;then
;Вывод сообщения об ошибке
				(alert Error_004_TXT)
;else
				(progn
;Перемещение курсора на вычисленную позицию
					(vlax-invoke-method CurRecordSet "Move" _KeyPosition 1)
;Загрузка строки
					(setq 	_Str (vlax-invoke-method CurRecordSet "GetString" 2 1 "\t" "\t" nil))
;Замена в строке повторяющихся последовател	"\t\t" на "\t \t"
					(repeat (length CurColumnList) (setq _Str (vl-string-subst "\t \t" "\t\t" _Str)))
;Очистка таблицы от записанных значений
					(dcl_Grid_Clear k410odcl_ChoiceRecordDialog_Grid)
;Запись в таблицу загруженной строки
					(dcl_Grid_AddString k410odcl_ChoiceRecordDialog_Grid _Str)
;Активация кнопки "Ok"			
					(dcl_Control_SetProperty k410odcl_ChoiceRecordDialog_OkButton "Enabled" T)
				)
			)
		)
	)
)
;--------------------------------------------------------- Заполнение таблицы --

;---------------------------------------- Инициализация диалога выбора записи --
(defun c:k410odcl_ChoiceRecordDialog_OnInitialize (/)
;Инициализация переменных	
	(InitBlock)
;Запись значений идентификации источника данных
	(dcl_Control_SetProperty k410odcl_ChoiceRecordDialog_LabelUDLFilePrompt "Caption" TagUDLFile)
	(dcl_Control_SetProperty k410odcl_ChoiceRecordDialog_LabelTableNamePrompt "Caption" TagTableName)
	(dcl_Control_SetProperty k410odcl_ChoiceRecordDialog_LabelKeyFieldPrompt "Caption" TagKeyFieldName)
;Сброс элементов управления
	(dcl_Control_SetProperty k410odcl_ChoiceRecordDialog_SearchButton "Enabled" F)
	(dcl_Control_SetProperty k410odcl_ChoiceRecordDialog_OkButton "Enabled" F)
;Загрузка набора данных
	(GetRecordSet)
;Заполнение таблицы диалога выбора записи
	(FillChoiceRecordDialogTable)
;Сброс счетчика обновлений
	(setq ItemIndex 0)
;Установка фокуса на поле вводв	
	(dcl_Control_SetFocus k410odcl_ChoiceRecordDialog_KeyFieldTextBox)
)
;---------------------------------------- Инициализация диалога выбора записи --

;---------------------------------------------- Смена значения ключевого поля --
(defun c:k410odcl_ChoiceRecordDialog_KeyFieldTextBox_OnEditChanged (NewValue /)
;Активация кнопки "Найти"
	(dcl_Control_SetProperty k410odcl_ChoiceRecordDialog_SearchButton "Enabled" T)
;Сброс кнопки "Ok"			
	(dcl_Control_SetProperty k410odcl_ChoiceRecordDialog_OkButton "Enabled" F) 
)
;---------------------------------------------- Смена значения ключевого поля --

;----------------------------------------------------------- Кнопка "Закрыть" --
(defun c:k410odcl_ChoiceRecordDialog_CancelButton_OnClicked (/)
	(dcl_Form_Close k410odcl_ChoiceRecordDialog)
)
;----------------------------------------------------------- Кнопка "Закрыть" --

;--------------------------------------------------------------- Поиск записи --
(defun RecordSearch (/ 	_KeyPosition
						_Str)
;Вычисление позиции ключевого поля в списке значений ключевого поля
	(setq _KeyPosition (vl-position (dcl_Control_GetProperty k410odcl_ChoiceRecordDialog_KeyFieldTextBox "Text") CurKeyValueList))
	(if (null _KeyPosition)
;then
;Вывод сообщения об ошибке
		(alert Error_004_TXT)
;else
		(progn
;Перемещение курсора на вычисленную позицию
			(vlax-invoke-method CurRecordSet "Move" _KeyPosition 1)
;Загрузка строки
			(setq 	_Str (vlax-invoke-method CurRecordSet "GetString" 2 1 "\t" "\t" nil))
;Замена в строке повторяющихся последовател	"\t\t" на "\t \t"
			(repeat (length CurColumnList) (setq _Str (vl-string-subst "\t \t" "\t\t" _Str)))
;Очистка таблицы от записанных значений
			(dcl_Grid_Clear k410odcl_ChoiceRecordDialog_Grid)
;Запись в таблицу загруженной строки
			(dcl_Grid_AddString k410odcl_ChoiceRecordDialog_Grid _Str)
;Сброс кнопки "Найти"
			(dcl_Control_SetProperty k410odcl_ChoiceRecordDialog_SearchButton "Enabled" F)
;Активация кнопки "Ok"			
			(dcl_Control_SetProperty k410odcl_ChoiceRecordDialog_OkButton "Enabled" T)
		)
	)
)
;--------------------------------------------------------------- Поиск записи --

;------------------------------------------------------------- Кнопка "Найти" --
(defun c:k410odcl_ChoiceRecordDialog_SearchButton_OnClicked (/)
	(RecordSearch)
)
;------------------------------------------------------------- Кнопка "Найти" --

;------------------------------------------------- Запись значений в атрибуты --
(defun WriteToAttrs (/	_AttrIndex
						_ColPos
						_Layers)
;Слои документа
	(setq _Layers (vla-get-Layers ActiveDoc))
;Запись значений в атрибуты из таблицы
	(foreach _AttrRef (vlax-safearray->list (vlax-variant-value (vla-GetAttributes CurBlockRef)))
;Поиск описания атрибута
		(setq _AttrIndex (vl-position (vla-Get-TagString _AttrRef) AttrList))
;Запись значения в атрибут
		(if (not (null _AttrIndex))
;then
			(progn
;Определение смещения в списке колонок
				(setq _ColPos (vl-position (nth _AttrIndex AttrFieldsList) CurColumnList))
				(if (= (vla-get-Lock (vla-Item _Layers (vla-get-Layer _AttrRef))) :vlax-true)
					(prompt "Атрибут на блокированном слое\n")	;then
;Запись значения в атрибут
					(progn	;else
						(if (not (null _ColPos))
;then				
							(vla-Put-TextString _AttrRef (dcl_Grid_GetCellText k410odcl_ChoiceRecordDialog_Grid 0 _ColPos))
;else
							(if (= (dcl_Control_GetProperty k410odcl_ChoiceRecordDialog_CheckBox2 "Value") 1)
								(vla-Put-TextString _AttrRef "")
							)
						)
;Запись в атрибут значения ключевого поля
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
;------------------------------------------------- Запись значений в атрибуты --

;---------------------------------------------------------------- Кнопка "Ok" --
(defun c:k410odcl_ChoiceRecordDialog_OkButton_OnClicked (/)
	(WriteToAttrs)
;Отключение кнопки
	(dcl_Control_SetProperty k410odcl_ChoiceRecordDialog_OkButton "Enabled" F)
;Закрытие диалога
	(if (= (dcl_Control_GetProperty k410odcl_ChoiceRecordDialog_CheckBox1 "Value") 1)
		(dcl_Form_Close k410odcl_ChoiceRecordDialog)
	)
)
;---------------------------------------------------------------- Кнопка "Ok" --

;--------------------------------------------- Кнопка "Обновить набор данных" --
(defun c:k410odcl_ChoiceRecodDialog_ReloadDSButton_OnClicked (/)
;Обновиить набор данных
	(ReloadRecordSet)
;Заполнение таблицы диалога выбора записи
	(FillChoiceRecordDialogTable)
)
;--------------------------------------------- Кнопка "Обновить набор данных" --

;--------------------------------------------------- Кнопка "Обновить данные" --
(defun c:k410odcl_ChoiceRecordDialog_ReloadDataButton_OnClicked (/)
;Установка флага диалога обновления
	(setq ReloadFlag ReloadData)
;Диалог обновления набора данных
	(prompt "\n")
;Отображение формы
	(dcl_Form_Show k410odcl_ReloadDialog)
;Сброс флага диалога обновления
	(setq ReloadFlag 0)
)
;--------------------------------------------------- Кнопка "Обновить данные" --

;----------------------------------- CheckBox "Заполнять атрибуты при ошибке" --
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
;----------------------------------- CheckBox "Заполнять атрибуты при ошибке" --

;------------------------------------------------------------- Кнопка "Enter" --
(defun c:k410odcl_ChoiceRecordDialog_OnOK (/)
  (RecordSearch)
  (WriteToAttrs)
)
;------------------------------------------------------------- Кнопка "Enter" -- 
;####################################################### Диалог: выбор записи ##

;################################################### Диалог: массив вхождений ##
(defun BlockRefArray (/)
	(prompt "\n")
	(if (= (vla-get-ActiveSpace ActiveDoc) 1)
		(dcl_Form_Show k410odcl_BlockRefArray);then
		(alert "Функция реализована только для объектов пространства модели");else
	)
	(princ)
)


;------------------------- Инициализация диалога построения массива вхождений --
(defun c:k410odcl_BlockRefArray_OnInitialize (/)
;Заполнение полей элементов управления
	(if (= (dcl_Control_GetProperty k410odcl_BlockRefArray_WidthTextBox "Text") "")
		(dcl_Control_SetProperty k410odcl_BlockRefArray_WidthTextBox "Text" ArrayWidth))
		
	(if (= (dcl_Control_GetProperty k410odcl_BlockRefArray_dYTextBox "Text") "")
		(dcl_Control_SetProperty k410odcl_BlockRefArray_dYTextBox "Text" ArraydX))
		
	(if (= (dcl_Control_GetProperty k410odcl_BlockRefArray_dXTextBox "Text") "")
		(dcl_Control_SetProperty k410odcl_BlockRefArray_dXTextBox "Text" ArraydY))

;Сброс элементов управления
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
;Инициализация блока
	(InitBlock)
;Загрузка набора данных
	(GetRecordSet)
;Установка элементов управления
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

;------------------------- Инициализация диалога построения массива вхождений --

;------------------------------------------------- Создание массива вхождений --
(defun CreateBlockRefArray (/	_KeyValueList
								_MSpace
								_Width
								_dX
								_dY
								_InsPoint
								_BlockName
								_BasePoint
								_Index)
;Загрузка настроечных переменных массива
	(setq	_MSpace (vla-get-ModelSpace ActiveDoc)
			_Width (atoi (dcl_Control_GetProperty k410odcl_BlockRefArray_WidthTextBox "Text"))
			_dX (atoi (dcl_Control_GetProperty k410odcl_BlockRefArray_dXTextBox "Text"))
			_dY (atoi (dcl_Control_GetProperty k410odcl_BlockRefArray_dYTextBox "Text"))
			_BlockName (vla-get-EffectiveName CurBlockRef)
			ReloadDataFlag T
			_Index 0)
;Загрузка списка значений ключей
	(if (= (dcl_Control_GetProperty k410odcl_BlockRefArray_OptionButton1 "Value") 1)
		(setq _KeyValueList CurKeyValueList);then
		(setq _KeyValueList KeyValues);else
	)
;Закрытие окна диалога
	(dcl_Form_Close k410odcl_BlockRefArray)
;Выбор базовой точки массива
	(setq 	_InsPoint (getpoint)
			_BasePoint _InsPoint)
;Построение массива
	(foreach _KeyFieldValue _KeyValueList
;Добавление вхождения
		(setq 	CurBlockRef (vla-InsertBlock _MSpace (vlax-3d-point _InsPoint) _BlockName 1 1 1 0)
				_Index (1+ _Index))
;Инициализация созданного вхождения
		(InitBlock)
;Запись значения ключевого поля во вхождение
		(vla-Put-TextString AttrRefKeyFieldValue _KeyFieldValue)
;Расчет точки вставки следующего вхождения
		(if (= _Index _Width)
			(setq	_Index 0
					_InsPoint (list (car _BasePoint) (+ (cadr _InsPoint) _dY) 0))
			(setq _InsPoint (list (+ (car _InsPoint) _dX) (cadr _InsPoint) 0))
		)
	)
	(setq ReloadDataFlag F)
)
;------------------------------------------------- Создание массива вхождений --

;---------------------------------------------------------------- Кнопка "Ok" --
(defun c:k410odcl_BlockRefArray_OkButton_OnClicked (/)
	(CreateBlockRefArray)
)
;---------------------------------------------------------------- Кнопка "Ok" --

;----------------------------------------------------------- Кнопка "Закрыть" --
(defun c:k410odcl_BlockRefArray_CancelButton_OnClicked (/)
  (dcl_Form_Close k410odcl_BlockRefArray)
)

;----------------------------------------------------------- Кнопка "Закрыть" --

;--------------------------------------------- Кнопка "Обновить набор данных" --
(defun c:k410odcl_BlockRefArray_ReloadDSButton_OnClicked (/)
 ;Обновиить набор данных
	(ReloadRecordSet)
)
;--------------------------------------------- Кнопка "Обновить набор данных" --

;------------------------------------------------------------- Кнопка "Enter" --
(defun c:k410odcl_BlockRefArray_OnOK (/)
  (CreateBlockRefArray))
;------------------------------------------------------------- Кнопка "Enter" --

;------------------------------------------------- Контроль вводимых значений --
(defun c:k410odcl_BlockRefArray_WidthTextBox_OnEditChanged (NewValue /)
	(if (and	(<= (atoi NewValue) 0)
					(/= NewValue ""))
		(progn
			(alert "Количество рядов или столбцов должно быть больше 0")
			(dcl_Control_SetProperty k410odcl_BlockRefArray_WidthTextBox "Text" ArrayWidth)
		)
	)
)

(defun c:k410odcl_BlockRefArray_WidthTextBox_OnKillFocus (/)
	(if (= (dcl_Control_GetProperty k410odcl_BlockRefArray_WidthTextBox "Text") "")
		(progn
			(alert "Количество рядов или столбцов должно быть больше 0")
			(dcl_Control_SetProperty k410odcl_BlockRefArray_WidthTextBox "Text" ArrayWidth)
		)
	)
)
;-------------------------------------------------------------------------------
(defun c:k410odcl_BlockRefArray_dYTextBox_OnKillFocus (/)
	(if (= (dcl_Control_GetProperty k410odcl_BlockRefArray_dYTextBox "Text") "")
		(progn
			(alert "Неверное расстояние между рядами")
			(dcl_Control_SetProperty k410odcl_BlockRefArray_dYTextBox "Text" ArraydY)
		)
	)
)
;-------------------------------------------------------------------------------
(defun c:k410odcl_BlockRefArray_dXTextBox_OnKillFocus (/)
	(if (= (dcl_Control_GetProperty k410odcl_BlockRefArray_dXTextBox "Text") "")
		(progn
			(alert "Неверное расстояние между столбцами")
			(dcl_Control_SetProperty k410odcl_BlockRefArray_dXTextBox "Text" ArraydX)
		)
	)
)

;------------------------------------------------- Контроль вводимых значений --
;################################################### Диалог: массив вхождений ##

;######################################################### Диалог: обновление ##
;------------------------------------------------------ Инициализация диалога --
(defun c:k410odcl_ReloadDialog_OnInitialize (/)
;Обновление набора данных
	(if (= ReloadFlag ReloadRecordSet)
		(progn
;Заголовок окна
			(dcl_Control_SetProperty k410odcl_ReloadDialog "TitleBarText" (strcat TitleTXT "Обновить набор данных"))
;Перемещение полосы прогресса в начало
			(dcl_Control_SetProperty k410odcl_ReloadDialog_ProgressBar "Value" 0)
;Отображение строки диалога
			(dcl_Control_SetProperty k410odcl_ReloadDialog_Label "Caption" "Обновить набор данных?")
		)
	)
;Обновление данных атрибутов
	(if (= ReloadFlag ReloadData)
		(progn
;Заголовок окна
			(dcl_Control_SetProperty k410odcl_ReloadDialog "TitleBarText" (strcat TitleTXT "Обновить данные атрибутов"))
;Перемещение полосы прогресса в начало
			(dcl_Control_SetProperty k410odcl_ReloadDialog_ProgressBar "Value" 0)
;Отображение строки диалога
			(dcl_Control_SetProperty k410odcl_ReloadDialog_Label "Caption" "Обновить данные атрибутов?")
		)
	)
)

;------------------------------------------------------ Инициализация диалога --

;---------------------------------------------------------------- Кнопка "Да" --
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
;Сброс/активация кнопок
	(dcl_Control_SetProperty k410odcl_ReloadDialog_YesButton "Enabled" F)
	(dcl_Control_SetProperty k410odcl_ReloadDialog_NoButton "Enabled" F)
;Отображение полосы прокрутки
	(dcl_Control_SetProperty k410odcl_ReloadDialog_ProgressBar "Visible" T)
;Скрытыие метки
	(dcl_Control_SetProperty k410odcl_ReloadDialog_Label "Visible" F)
	(dcl_Control_Redraw k410odcl_ReloadDialog)
;Обновление набора данных
	(if (= ReloadFlag ReloadRecordSet)
		(progn
;Длина полосы прогресса (по количеству этапов)
			(dcl_Control_SetProperty k410odcl_ReloadDialog_ProgressBar "MaxValue" 2)
;Удаление из списка Views текущего набора данных	
			(setq Views (vl-remove (assoc (strcat TagUDLFile ";" TagTableName ";" TagKeyFieldName) Views) Views))
;Перемещение полосы прогресса
			(dcl_Control_SetProperty k410odcl_ReloadDialog_ProgressBar "Value" 1)
;Загрузка набора данных
			(GetRecordSet)
;Перемещение полосы прогресса
			(dcl_Control_SetProperty k410odcl_ReloadDialog_ProgressBar "Value" 2)
		)
	)
;Обновление данных атрибутов
	(if (= ReloadFlag ReloadData)
		(progn
;Имя обекта
			(setq 	_BlockRefObjectName "AcDbBlockReference"
;Сохранение текущего вхождения
					_CurRef CurBlockRef
;Имя описания текущего вхождения
					_CurAffectiveName (vla-get-EffectiveName CurBlockRef)
;Пространство модели
					_MSpace (vla-get-ModelSpace ActiveDoc)
;Загрузка количества элементов в пространстве модели			
					_MSItemCount (vla-get-count _MSpace)
;Сброс строки записи
					_Str ""
;Количество обработанных вхождений
					_ReloadCount 0
;Количество ошибок при обновлении
					_ErrorCount 0
;Количество вхождений на блокированных слоях
					_LockCount 0
;Слои документа
					_Layers (vla-get-Layers ActiveDoc)
;Установка флага обновления данных
					ReloadDataFlag T)
	
;Установка правой границы полосы прокрутки
			(dcl_Control_SetProperty k410odcl_ReloadDialog_ProgressBar "MaxValue" _MSItemCount)
;Перемещение полосы в начало
			(dcl_Control_SetProperty k410odcl_ReloadDialog_ProgressBar "Value" ItemIndex)
;Обновление элементов
			(while (and ReloadDataFlag (/= ItemIndex _MSItemCount))
;Загрузка ссылки на очередной элемент из пространства модели или простраства листа
				(setq _ReloadItem (vla-item _MSpace ItemIndex))
;Обновление, если элемент -- вхождение блока выбранного описания
				(if (and 	(= (vla-get-ObjectName _ReloadItem) _BlockRefObjectName)
							(= (vla-get-EffectiveName _ReloadItem) _CurAffectiveName))
;Проверка слоя на блокировку
					(if (= (vla-get-Lock (vla-Item _Layers (vla-get-Layer _ReloadItem))) :vlax-true)	;then
						(setq _LockCount (1+ _LockCount))	;then
						(progn	;else
;Запись текущего элемента пространства модели
							(setq 	CurBlockRef _ReloadItem
;Обнуление строки записи источника базы данных
									_Str "")
;Инициализация вхождения блока
							(InitBlock)
;Обработка значения ключевого поля
							(if (= TagKeyFieldValue "")
								(progn	;then
;Заполнение атрибутов
									(if (= (dcl_Control_GetProperty k410odcl_ChoiceRecordDialog_FillErrorCheckBox "Value") 1)
										(progn
											(repeat (length CurColumnList) 
												(setq 	_Str 	(strcat _Str 
																	(dcl_Control_GetProperty k410odcl_ChoiceRecordDialog_FillErrorTextBox "Text")
																	"\t")))
;Очистка таблицы от записанных значений
											(dcl_Grid_Clear k410odcl_ChoiceRecordDialog_Grid)
;Запись в таблицу сформированной строки
											(dcl_Grid_AddString k410odcl_ChoiceRecordDialog_Grid _Str)
;Запись значений в атрибуты
											(WriteToAttrs)
										)
									)
;Учет ошибки
									(setq _ErrorCount (1+ _ErrorCount))
								)
								(progn	;else
;Вычисление позиции ключевого поля в списке значений ключевого поля
									(setq _KeyPosition (vl-position TagKeyFieldValue CurKeyValueList))
									(if (null _KeyPosition)
										(progn	;then
;Заполнение атрибутов
											(if (= (dcl_Control_GetProperty k410odcl_ChoiceRecordDialog_FillErrorCheckBox "Value") 1)
												(progn
													(repeat (length CurColumnList) 
														(setq _Str (strcat 
																	_Str 
																	(dcl_Control_GetProperty k410odcl_ChoiceRecordDialog_FillErrorTextBox "Text")
																	"\t")))
;Очистка таблицы от записанных значений
													(dcl_Grid_Clear k410odcl_ChoiceRecordDialog_Grid)
;Запись в таблицу сформированной строки
													(dcl_Grid_AddString k410odcl_ChoiceRecordDialog_Grid _Str)
;Выполнение события от кнопки "Ok"
													(c:k410odcl_ChoiceRecordDialog_OkButton_OnClicked)
												)
											)
;Учет ошибки
											(setq _ErrorCount (1+ _ErrorCount))
										)
										(progn	;else
;Перемещение курсора на вычисленную позицию
											(vlax-invoke-method CurRecordSet "Move" _KeyPosition 1)
;Загрузка строки
											(setq 	_Str (vlax-invoke-method CurRecordSet "GetString" 2 1 "\t" "\t" nil))
;Замена в строке повторяющихся последовател	"\t\t" на "\t \t"
											(repeat (length CurColumnList) (setq _Str (vl-string-subst "\t \t" "\t\t" _Str)))
;Очистка таблицы от записанных значений
											(dcl_Grid_Clear k410odcl_ChoiceRecordDialog_Grid)
;Запись в таблицу загруженной строки
											(dcl_Grid_AddString k410odcl_ChoiceRecordDialog_Grid _Str)
;Выполнение события от кнопки "Ok"
											(c:k410odcl_ChoiceRecordDialog_OkButton_OnClicked)
										)
									)
								)
							)
							(setq _ReloadCount (1+ _ReloadCount))
						)
					)
				)
;Переход на следующий элемент
				(setq ItemIndex (1+ ItemIndex))
;Увеличинение значения полосы прокрутки
				(dcl_Control_SetProperty k410odcl_ReloadDialog_ProgressBar "Value" ItemIndex)
			)
;Сброс флага обновления данных
			(setq 	ReloadDataFlag F
;Восстанвление ссылки на исходное вхождение
					CurBlockRef _CurRef)
;Инициализация вхождения блока
			(InitBlock)
;Очистка таблицы от записанных значений
			(dcl_Grid_Clear k410odcl_ChoiceRecordDialog_Grid)
;Обработка значения ключевого поля
			(if (= TagKeyFieldValue "")
;then
				(dcl_Control_SetProperty k410odcl_ChoiceRecordDialog_KeyFieldTextBox "Text" KeyFieldValueTXT)
;else
				(progn
					(dcl_Control_SetProperty k410odcl_ChoiceRecordDialog_KeyFieldTextBox "Text" TagKeyFieldValue)
;Вычисление позиции ключевого поля в списке значений ключевого поля
					(setq _KeyPosition (vl-position TagKeyFieldValue CurKeyValueList))
					(if (null _KeyPosition)
;then
;Ошибка
						(princ)
;else
						(progn
;Перемещение курсора на вычисленную позицию
							(vlax-invoke-method CurRecordSet "Move" _KeyPosition 1)
;Загрузка строки
							(setq 	_Str (vlax-invoke-method CurRecordSet "GetString" 2 1 "\t" "\t" nil))
;Замена в строке повторяющихся последовател	"\t\t" на "\t \t"
							(repeat (length CurColumnList) (setq _Str (vl-string-subst "\t \t" "\t\t" _Str)))
;Очистка таблицы от записанных значений
							(dcl_Grid_Clear k410odcl_ChoiceRecordDialog_Grid)
;Запись в таблицу загруженной строки
							(dcl_Grid_AddString k410odcl_ChoiceRecordDialog_Grid _Str)
						)
					)
				)
			)
;Результат работы
			(alert 	(strcat "Объектов:		" (itoa ItemIndex) "\n"
							"Вхождений:		" (itoa (+ _ReloadCount _LockCount)) "\n"
							"Ошибок:			" (itoa _ErrorCount) "\n"
							"На блокированных слоях:	" (itoa _LockCount)))
;Сброс указателя элементов, если просмотрены все элементы документа
			(if (= ItemIndex _MSItemCount)
				(setq ItemIndex 0))
		)
	)
;Сброс/активация кнопок
	(dcl_Control_SetProperty k410odcl_ReloadDialog_YesButton "Enabled" T)
	(dcl_Control_SetProperty k410odcl_ReloadDialog_NoButton "Enabled" T)
;Скрытыие полосы прокрутки
	(dcl_Control_SetProperty k410odcl_ReloadDialog_ProgressBar "Visible" F)
;Отображение метки
	(dcl_Control_SetProperty k410odcl_ReloadDialog_Label "Visible" T)
;Закрыть диалог
	(dcl_Form_Close k410odcl_ReloadDialog)
)
;---------------------------------------------------------------- Кнопка "Да" --

;--------------------------------------------------------------- Кнопка "Нет" --
(defun c:k410odcl_ReloadDialog_NoButton_OnClicked (/)
;Закрыть диалог
	(dcl_Form_Close k410odcl_ReloadDialog)
)
;--------------------------------------------------------------- Кнопка "Нет" --
;######################################################### Диалог: обновление ##

(princ)
;http://www.askit.ru/custom/progr_admin/m13/13_01_ado_basics.htm