 
#Если Сервер Или ТолстыйКлиентОбычноеПриложение Или ВнешнееСоединение Тогда

//++ Проф-ИТ, #53, Башинская А., 13.10.2023	
	
#Область ПрограммныйИнтерфейс

Функция ЕстьПрофильДоступа(Пользователь, Профиль) Экспорт
	
	УстановитьПривилегированныйРежим(Истина);
	Запрос = Новый Запрос();
	Запрос.Текст = 
	"ВЫБРАТЬ
	|	ГруппыДоступа.Ссылка КАК Ссылка,
	|	ГруппыДоступа.Профиль КАК Профиль
	|ИЗ
	|	Справочник.ГруппыДоступа КАК ГруппыДоступа
	|ГДЕ
	|	ГруппыДоступа.Пользователи.Пользователь = &Пользователь
	|	И ГруппыДоступа.Профиль = &Профиль";
	Запрос.УстановитьПараметр("Пользователь", Пользователь);
	Запрос.УстановитьПараметр("Профиль", Профиль);
	Выборка = Запрос.Выполнить().Выбрать();
	
	Возврат Выборка.Следующий();
	
КонецФункции 

Процедура ЗагрузитьДанныеВДокумент(ВводНачальныхДанных, ТаблицаДанных, СообщениеРезультатЗарузки) Экспорт
		
	ПроверитьСоставТаблицыПронумеровать(ТаблицаДанных, СообщениеРезультатЗарузки);     
	Если ЗначениеЗаполнено(СообщениеРезультатЗарузки) Тогда
		Возврат;
	Иначе             
		ВыборкаДанные = ДополнитьДанныеИзБД(ТаблицаДанных, СообщениеРезультатЗарузки);
		ЗагрузитьДанныеВОбъект(ВводНачальныхДанных, ВыборкаДанные, СообщениеРезультатЗарузки);
	КонецЕсли;
	
КонецПроцедуры

#КонецОбласти
	
#Область СлужебныеПроцедурыИФункции

Процедура ПроверитьСоставТаблицыПронумеровать(ТаблицаДанных, СообщениеРезультатЗарузки) 
	
	Если ТаблицаДанных.Колонки.Найти("ОбъектРемонта") = Неопределено Тогда
		СообщениеРезультатЗарузки = СообщениеРезультатЗарузки + "Не найдена колонка ""Объект ремонта"";" + Символы.ПС; 
	КонецЕсли;
	Если ТаблицаДанных.Колонки.Найти("Инвентарный_") = Неопределено Тогда
		СообщениеРезультатЗарузки = СообщениеРезультатЗарузки + "Не найдена колонка ""Инвентарный №"";" + Символы.ПС; 
	КонецЕсли;
	Если ТаблицаДанных.Колонки.Найти("РаспространятьНаПодчиненныеОР") = Неопределено Тогда
		СообщениеРезультатЗарузки = СообщениеРезультатЗарузки
			+ "Не найдена колонка ""Распространять на подчиненные ОР"";" + Символы.ПС; 
	КонецЕсли;
	Если ТаблицаДанных.Колонки.Найти("СтруктураИерархии") = Неопределено Тогда
		СообщениеРезультатЗарузки = СообщениеРезультатЗарузки + "Не найдена колонка ""Структура иерархии"";" + Символы.ПС; 
	КонецЕсли;
	Если ТаблицаДанных.Колонки.Найти("ВидЭксплуатации") = Неопределено Тогда
		СообщениеРезультатЗарузки = СообщениеРезультатЗарузки + "Не найдена колонка ""Вид эксплуатации"";" + Символы.ПС; 
	КонецЕсли;
	Если ТаблицаДанных.Колонки.Найти("ВидРемонтныхРабот") = Неопределено Тогда
		СообщениеРезультатЗарузки = СообщениеРезультатЗарузки + "Не найдена колонка ""Вид ремонтных работ"";" + Символы.ПС; 
	КонецЕсли;	
	Если ТаблицаДанных.Колонки.Найти("ДатаОкончанияРемонта") = Неопределено Тогда
		СообщениеРезультатЗарузки = СообщениеРезультатЗарузки + "Не найдена колонка ""Дата окончания ремонта"";" + Символы.ПС; 
	КонецЕсли;
	Если ТаблицаДанных.Колонки.Найти("НаработкаПоказатель") = Неопределено Тогда
		СообщениеРезультатЗарузки = СообщениеРезультатЗарузки + "Не найдена колонка ""Наработка показатель"";" + Символы.ПС; 
	КонецЕсли;
	Если ТаблицаДанных.Колонки.Найти("НаработкаДата_работы_по") = Неопределено Тогда
		СообщениеРезультатЗарузки = СообщениеРезультатЗарузки
			+ "Не найдена колонка ""Наработка дата работы по"";" + Символы.ПС;
	КонецЕсли;
	Если ТаблицаДанных.Колонки.Найти("НаработкаЗначение") = Неопределено Тогда
		СообщениеРезультатЗарузки = СообщениеРезультатЗарузки + "Не найдена колонка ""Наработка значение"";" + Символы.ПС; 
	КонецЕсли;
	
	Если ЗначениеЗаполнено(СообщениеРезультатЗарузки) Тогда   
		СообщениеРезультатЗарузки = СообщениеРезультатЗарузки + "ЗАГРУЗКА ПРЕРВАНА." + Символы.ПС;
		Возврат;
	КонецЕсли;
	
	// Удалим пустые строки или строки без заполненного объекта ремонта
	сч = ТаблицаДанных.Количество() - 1;
	Пока сч >= 0 Цикл
		Стр = ТаблицаДанных[сч];
		Если СокрЛП(Стр.ОбъектРемонта) = "" Тогда 
			ТаблицаДанных.Удалить(Стр);
		КонецЕсли;		
		сч = сч - 1;
	КонецЦикла;

	// Пронумеруем и нормализуем
	ТаблицаДанных.Колонки.Добавить("НомерСтроки", Новый ОписаниеТипов("Число", Новый КвалификаторыЧисла(3, 0)));
	
	сч = 1;
	Для Каждого Стр Из ТаблицаДанных Цикл 
		Стр.ОбъектРемонта 		= СокрЛП(Стр.ОбъектРемонта); 
		Стр.Инвентарный_ 		= СокрЛП(Стр.Инвентарный_);
		Стр.СтруктураИерархии 	= СокрЛП(Стр.СтруктураИерархии);
		Стр.ВидЭксплуатации 	= СокрЛП(Стр.ВидЭксплуатации);
		Стр.ВидРемонтныхРабот 	= СокрЛП(Стр.ВидРемонтныхРабот);
		Стр.НаработкаПоказатель = СокрЛП(Стр.НаработкаПоказатель);
		Стр.НаработкаЗначение   = СокрЛП(Стр.НаработкаЗначение);
			
		Стр.РаспространятьНаПодчиненныеОР = СокрЛП(НРег(Стр.РаспространятьНаПодчиненныеОР));
		Стр.ДатаОкончанияРемонта 		  = СокрЛП(Стр.ДатаОкончанияРемонта);      
		Стр.НаработкаДата_работы_по 	  = СокрЛП(Стр.НаработкаДата_работы_по);

		Стр.НомерСтроки = сч;
		сч = сч + 1;
	КонецЦикла;  
	
КонецПроцедуры

Процедура ЗагрузитьДанныеВОбъект(ВводНачальныхДанных, ВыборкаПоОР, СообщениеРезультатЗарузки) 
	
	// Очистим ТЧ
	ВводНачальныхДанных.ОбъектыРемонта.Очистить();
	ВводНачальныхДанных.ВыполненныеРаботы.Очистить();
	ВводНачальныхДанных.НаработкаОбъектыРемонта.Очистить();
	
	СообщениеРезультатЗарузки = СообщениеРезультатЗарузки + "ЗАГРУЗКА ДАННЫХ В ДОКУМЕНТ." + Символы.ПС;
	
	СтруктураОтбора = Новый Структура("ОбъектРемонта");
	
	Пока ВыборкаПоОР.Следующий() Цикл
		
		Если ВыборкаПоОР.СуществущийДокументПоОР <> Null Тогда				
			Продолжить;
		КонецЕсли;
		
		ВыборкаДанные = ВыборкаПоОР.Выбрать();    
	
		Пока ВыборкаДанные.Следующий() Цикл
		
			Если ВыборкаДанные.Пропустить = Истина Тогда
				Продолжить;
			КонецЕсли;
			СтруктураОтбора.ОбъектРемонта = ВыборкаДанные.ОбъектРемонта;
				
			ЗаполнитьТЧОбъектыРемонта(ВыборкаДанные, ВводНачальныхДанных, СообщениеРезультатЗарузки, СтруктураОтбора);
			ЗаполнитьТЧВыполненныеРаботы(ВыборкаДанные, ВводНачальныхДанных, СообщениеРезультатЗарузки);
			ЗаполнитьТЧНаработкаОбъектыРемонта(ВыборкаДанные, ВводНачальныхДанных, СообщениеРезультатЗарузки);
		  				
		КонецЦикла;
		
	КонецЦикла;   

	СообщениеРезультатЗарузки = СообщениеРезультатЗарузки + Символы.ПС + "ЗАГРУЗКА ЗАВЕРШЕНА." + Символы.ПС;
		
КонецПроцедуры

Процедура ЗаполнитьТЧОбъектыРемонта(ВыборкаДанные, ВводНачальныхДанных, СообщениеРезультатЗарузки, СтруктураОтбора)
	
	
	Если ЗначениеЗаполнено(ВыборкаДанные.СтруктураИерархии)
	Или ЗначениеЗаполнено(ВыборкаДанные.ВидЭксплуатации) Тогда				
		// Проверим, не добавлены ли уже данные по ОР в ТЧ - тогда не загружаем
		Если ВводНачальныхДанных.ОбъектыРемонта.НайтиСтроки(СтруктураОтбора).Количество() > 0 Тогда  
			// В ТЧ уже есть строка по этому объекту ремонта, повторно не добавляем
			//СообщениеРезультатЗарузки = СтрШаблон("%1В таблицу ""Объекты ремонта"" по объекту ремонта %2"
			//									  + " данные уже добавлены ранее. По строке №%3 данные не загружены.%4", 
			//						 			  	СообщениеРезультатЗарузки, ВыборкаДанные.ОбъектРемонта, ВыборкаДанные.НомерСтроки, Символы.ПС);    
		Иначе
			НСтр 								= ВводНачальныхДанных.ОбъектыРемонта.Добавить();
			НСтр.ОбъектРемонта 					= ВыборкаДанные.ОбъектРемонта;
			НСтр.СтруктураИерархии 				= ВыборкаДанные.СтруктураИерархии;
			НСтр.ТекущееСостояние 				= ВыборкаДанные.ВидЭксплуатации;
			НСтр.ИзменятьСостояниеПодчиненныхОР = ВыборкаДанные.РаспространятьНаПодчиненныеОР;		
			
			СообщениеРезультатЗарузки = СтрШаблон("%1В таблицу ""Объекты ремонта"" загружена строка %2.%3", 
									 			  СообщениеРезультатЗарузки, ВыборкаДанные.НомерСтроки, Символы.ПС);    
		КонецЕсли;
	КонецЕсли;   	
	
КонецПроцедуры

Процедура ЗаполнитьТЧВыполненныеРаботы(ВыборкаДанные, ВводНачальныхДанных, СообщениеРезультатЗарузки)

	Если ЗначениеЗаполнено(ВыборкаДанные.ВидРемонтныхРабот)
	Или ЗначениеЗаполнено(ВыборкаДанные.ДатаОкончанияРемонта) Тогда				
		//// Проверим, не добавлены ли уже данные по ОР в ТЧ - тогда не загружаем
		//Если ВводНачальныхДанных.ВыполненныеРаботы.НайтиСтроки(СтруктураОтбора).Количество() > 0 Тогда
		//	СообщениеРезультатЗарузки = СтрШаблон("%1В таблицу ""Выполненные работы"" по объекту ремонта %2"
		//										  + " данные уже добавлены ранее. По строке №%3 данные не загружены.%4", 
												//СообщениеРезультатЗарузки, ВыборкаДанные.ОбъектРемонта, ВыборкаДанные.НомерСтроки, Символы.ПС);
		//Иначе
			НСтр 					  = ВводНачальныхДанных.ВыполненныеРаботы.Добавить();
			НСтр.ОбъектРемонта 		  = ВыборкаДанные.ОбъектРемонта;
			НСтр.ВидРемонтныхРабот 	  = ВыборкаДанные.ВидРемонтныхРабот;
			НСтр.ДатаОкончанияРемонта = СтрокаВДату(ВыборкаДанные.ДатаОкончанияРемонта);
			
			СообщениеРезультатЗарузки = СтрШаблон("%1В таблицу ""Выполненные работы"" загружена строка %2.%3", 
									 			  СообщениеРезультатЗарузки, ВыборкаДанные.НомерСтроки, Символы.ПС);    	
		//КонецЕсли;
	КонецЕсли; 	
	
КонецПроцедуры

Процедура ЗаполнитьТЧНаработкаОбъектыРемонта(ВыборкаДанные, ВводНачальныхДанных, СообщениеРезультатЗарузки)

	Если ЗначениеЗаполнено(ВыборкаДанные.НаработкаДата_работы_по)
	Или ЗначениеЗаполнено(ВыборкаДанные.НаработкаПоказатель)
	Или ЗначениеЗаполнено(ВыборкаДанные.НаработкаЗначение) Тогда
		//// Проверим, не добавлены ли уже данные по ОР в ТЧ - тогда не загружаем
		//Если ВводНачальныхДанных.НаработкаОбъектыРемонта.НайтиСтроки(СтруктураОтбора).Количество() > 0 Тогда
		//	СообщениеРезультатЗарузки = СтрШаблон("%1В таблицу ""Наработка объекты ремонта"" по объекту ремонта %2"
		//										  + " данные уже добавлены ранее. По строке №%3 данные не загружены.%4", 
												//СообщениеРезультатЗарузки, ВыборкаДанные.ОбъектРемонта, ВыборкаДанные.НомерСтроки, Символы.ПС);
		//Иначе
			НСтр 			   = ВводНачальныхДанных.НаработкаОбъектыРемонта.Добавить();
			НСтр.ОбъектРемонта = ВыборкаДанные.ОбъектРемонта;
			НСтр.Показатель    = ВыборкаДанные.НаработкаПоказатель;
			НСтр.ДатаРаботыПо  = СтрокаВДату(ВыборкаДанные.НаработкаДата_работы_по);
			НСтр.Значение 	   = ВыборкаДанные.НаработкаЗначение;		
			
			СообщениеРезультатЗарузки = СтрШаблон("%1В таблицу ""Наработка объекты ремонта"" загружена строка %2.%3", 
									 			  СообщениеРезультатЗарузки, ВыборкаДанные.НомерСтроки, Символы.ПС);    	
		//КонецЕсли;
	КонецЕсли;  	
	
КонецПроцедуры	

Функция ДополнитьДанныеИзБД(ТаблицаДанных, СообщениеРезультатЗарузки)        
	
	СообщениеРезультатЗарузки = СообщениеРезультатЗарузки + "АНАЛИЗ ДАННЫХ ФАЙЛА." + Символы.ПС;

	Запрос = Новый Запрос;
	Запрос.Текст = "ВЫБРАТЬ
	               |	ВнешнийИсточник.НомерСтроки КАК НомерСтроки,
	               |	ВнешнийИсточник.ОбъектРемонта КАК ОбъектРемонта,
	               |	ВнешнийИсточник.Инвентарный_ КАК Инвентарный_,
	               |	ВЫБОР
	               |		КОГДА ВнешнийИсточник.РаспространятьНаПодчиненныеОР = ""да""
	               |			ТОГДА ИСТИНА
	               |		КОГДА ВнешнийИсточник.РаспространятьНаПодчиненныеОР = ""нет""
	               |			ТОГДА ЛОЖЬ
	               |		ИНАЧЕ ЛОЖЬ
	               |	КОНЕЦ КАК РаспространятьНаПодчиненныеОР,
	               |	ВнешнийИсточник.СтруктураИерархии КАК СтруктураИерархии,
	               |	ВнешнийИсточник.ВидЭксплуатации КАК ВидЭксплуатации,
	               |	ВнешнийИсточник.ВидРемонтныхРабот КАК ВидРемонтныхРабот,
	               |	ВнешнийИсточник.ДатаОкончанияРемонта КАК ДатаОкончанияРемонта,
	               |	ВнешнийИсточник.НаработкаПоказатель КАК НаработкаПоказатель,
	               |	ВнешнийИсточник.НаработкаДата_работы_по КАК НаработкаДата_работы_по,
	               |	ВнешнийИсточник.НаработкаЗначение КАК НаработкаЗначение
	               |ПОМЕСТИТЬ ВнешнийИсточник
	               |ИЗ
	               |	&ТЗ КАК ВнешнийИсточник
	               |;
	               |
	               |////////////////////////////////////////////////////////////////////////////////
	               |ВЫБРАТЬ
	               |	ВнешнийИсточник.НомерСтроки КАК НомерСтроки,
	               |	ВнешнийИсточник.РаспространятьНаПодчиненныеОР КАК РаспространятьНаПодчиненныеОР,
	               |	ВнешнийИсточник.ДатаОкончанияРемонта КАК ДатаОкончанияРемонта,
	               |	ВнешнийИсточник.НаработкаДата_работы_по КАК НаработкаДата_работы_по,
	               |	ВнешнийИсточник.НаработкаЗначение КАК НаработкаЗначение,
	               |	КОЛИЧЕСТВО(РАЗЛИЧНЫЕ ЕСТЬNULL(торо_ОбъектыРемонта.Ссылка, НЕОПРЕДЕЛЕНО)) КАК ОбъектРемонтаКолВо,
	               |	МАКСИМУМ(ЕСТЬNULL(торо_ОбъектыРемонта.Ссылка, НЕОПРЕДЕЛЕНО)) КАК ОбъектРемонта,
	               |	КОЛИЧЕСТВО(РАЗЛИЧНЫЕ ЕСТЬNULL(торо_СтруктурыОР.Ссылка, НЕОПРЕДЕЛЕНО)) КАК СтруктураИерархииКолВо,
	               |	МАКСИМУМ(ЕСТЬNULL(торо_СтруктурыОР.Ссылка, НЕОПРЕДЕЛЕНО)) КАК СтруктураИерархии,
	               |	КОЛИЧЕСТВО(РАЗЛИЧНЫЕ ЕСТЬNULL(торо_ВидыЭксплуатации.Ссылка, НЕОПРЕДЕЛЕНО)) КАК ВидЭксплуатацииКолВо,
	               |	МАКСИМУМ(ЕСТЬNULL(торо_ВидыЭксплуатации.Ссылка, НЕОПРЕДЕЛЕНО)) КАК ВидЭксплуатации,
	               |	КОЛИЧЕСТВО(РАЗЛИЧНЫЕ ЕСТЬNULL(торо_ВидыРемонтов.Ссылка, НЕОПРЕДЕЛЕНО)) КАК ВидРемонтныхРаботКолВо,
	               |	МАКСИМУМ(ЕСТЬNULL(торо_ВидыРемонтов.Ссылка, НЕОПРЕДЕЛЕНО)) КАК ВидРемонтныхРабот,
	               |	КОЛИЧЕСТВО(РАЗЛИЧНЫЕ ЕСТЬNULL(ПараметрыВыработкиОС.Ссылка, НЕОПРЕДЕЛЕНО)) КАК НаработкаПоказательКолВо,
	               |	МАКСИМУМ(ЕСТЬNULL(ПараметрыВыработкиОС.Ссылка, НЕОПРЕДЕЛЕНО)) КАК НаработкаПоказатель,
	               |	ЕСТЬNULL(торо_ОбъектыРемонта.Ссылка, ИСТИНА) КАК Пропустить
	               |ПОМЕСТИТЬ ВТ
	               |ИЗ
	               |	ВнешнийИсточник КАК ВнешнийИсточник
	               |		ЛЕВОЕ СОЕДИНЕНИЕ Справочник.торо_ОбъектыРемонта КАК торо_ОбъектыРемонта
	               |		ПО ВнешнийИсточник.ОбъектРемонта = торо_ОбъектыРемонта.Наименование
	               |			И ВнешнийИсточник.Инвентарный_ = торо_ОбъектыРемонта.ИнвентарныйНомер
	               |		ЛЕВОЕ СОЕДИНЕНИЕ Справочник.торо_СтруктурыОР КАК торо_СтруктурыОР
	               |		ПО ВнешнийИсточник.СтруктураИерархии = торо_СтруктурыОР.Наименование
	               |		ЛЕВОЕ СОЕДИНЕНИЕ Справочник.торо_ВидыЭксплуатации КАК торо_ВидыЭксплуатации
	               |		ПО ВнешнийИсточник.ВидЭксплуатации = торо_ВидыЭксплуатации.Наименование
	               |		ЛЕВОЕ СОЕДИНЕНИЕ Справочник.торо_ВидыРемонтов КАК торо_ВидыРемонтов
	               |		ПО ВнешнийИсточник.ВидРемонтныхРабот = торо_ВидыРемонтов.Наименование
	               |		ЛЕВОЕ СОЕДИНЕНИЕ Справочник.ПараметрыВыработкиОС КАК ПараметрыВыработкиОС
	               |		ПО ВнешнийИсточник.НаработкаПоказатель = ПараметрыВыработкиОС.Наименование
	               |
	               |СГРУППИРОВАТЬ ПО
	               |	ВнешнийИсточник.НомерСтроки,
	               |	ВнешнийИсточник.РаспространятьНаПодчиненныеОР,
	               |	ВнешнийИсточник.ДатаОкончанияРемонта,
	               |	ВнешнийИсточник.НаработкаЗначение,
	               |	ВнешнийИсточник.НаработкаДата_работы_по,
	               |	ЕСТЬNULL(торо_ОбъектыРемонта.Ссылка, ИСТИНА)
	               |;
	               |
	               |////////////////////////////////////////////////////////////////////////////////
	               |ВЫБРАТЬ
	               |	ВТ.НомерСтроки КАК НомерСтроки,
	               |	ВТ.РаспространятьНаПодчиненныеОР КАК РаспространятьНаПодчиненныеОР,
	               |	ВТ.ДатаОкончанияРемонта КАК ДатаОкончанияРемонта,
	               |	ВТ.НаработкаДата_работы_по КАК НаработкаДата_работы_по,
	               |	ВТ.НаработкаЗначение КАК НаработкаЗначение,
	               |	ВТ.ОбъектРемонтаКолВо КАК ОбъектРемонтаКолВо,
	               |	ВТ.ОбъектРемонта КАК ОбъектРемонта,
	               |	ВТ.СтруктураИерархииКолВо КАК СтруктураИерархииКолВо,
	               |	ВТ.СтруктураИерархии КАК СтруктураИерархии,
	               |	ВТ.ВидЭксплуатацииКолВо КАК ВидЭксплуатацииКолВо,
	               |	ВТ.ВидЭксплуатации КАК ВидЭксплуатации,
	               |	ВТ.ВидРемонтныхРаботКолВо КАК ВидРемонтныхРаботКолВо,
	               |	ВТ.ВидРемонтныхРабот КАК ВидРемонтныхРабот,
	               |	ВТ.НаработкаПоказательКолВо КАК НаработкаПоказательКолВо,
	               |	ВТ.НаработкаПоказатель КАК НаработкаПоказатель,
	               |	ВТ.Пропустить КАК Пропустить,
	               |	МАКСИМУМ(торо_ВводНачальныхДанныхОбъектыРемонта.Ссылка) КАК СуществущийДокументПоОР
	               |ИЗ
	               |	ВТ КАК ВТ
	               |		ЛЕВОЕ СОЕДИНЕНИЕ Документ.торо_ВводНачальныхДанных.ОбъектыРемонта КАК торо_ВводНачальныхДанныхОбъектыРемонта
	               |		ПО (ВТ.ОбъектРемонта = торо_ВводНачальныхДанныхОбъектыРемонта.ОбъектРемонта
	               |				И НЕ торо_ВводНачальныхДанныхОбъектыРемонта.Ссылка.ПометкаУдаления)
	               |
	               |СГРУППИРОВАТЬ ПО
	               |	ВТ.РаспространятьНаПодчиненныеОР,
	               |	ВТ.НаработкаЗначение,
	               |	ВТ.ОбъектРемонта,
	               |	ВТ.СтруктураИерархии,
	               |	ВТ.ВидЭксплуатации,
	               |	ВТ.ВидРемонтныхРабот,
	               |	ВТ.НаработкаПоказатель,
	               |	ВТ.Пропустить,
	               |	ВТ.НаработкаДата_работы_по,
	               |	ВТ.ДатаОкончанияРемонта,
	               |	ВТ.НомерСтроки,
	               |	ВТ.ОбъектРемонтаКолВо,
	               |	ВТ.СтруктураИерархииКолВо,
	               |	ВТ.ВидЭксплуатацииКолВо,
	               |	ВТ.ВидРемонтныхРаботКолВо,
	               |	ВТ.НаработкаПоказательКолВо
	               |ИТОГИ
	               |	МАКСИМУМ(СуществущийДокументПоОР)
	               |ПО
	               |	ОбъектРемонта";

	Запрос.УстановитьПараметр("ТЗ", ТаблицаДанных);
	ВыборкаПоОР = Запрос.Выполнить().Выбрать(ОбходРезультатаЗапроса.ПоГруппировкам);
	
	Пока ВыборкаПоОР.Следующий() Цикл
		
		Если ВыборкаПоОР.СуществущийДокументПоОР <> Null Тогда				
			СообщениеРезультатЗарузки = СтрШаблон("%1По объекту ремонта №%2 найден уже созданный документ с данными:" +
			" %3. Данные не загружены.%4", СообщениеРезультатЗарузки, ВыборкаПоОР.ОбъектРемонта, ВыборкаПоОР.СуществущийДокументПоОР, Символы.ПС);     
			Продолжить;
		КонецЕсли;
		
		Выборка = ВыборкаПоОР.Выбрать();    
	
		Пока Выборка.Следующий() Цикл		
			ПроверитьЗаполненностьДанных(Выборка, СообщениеРезультатЗарузки);						
		КонецЦикла;	 
		
	КонецЦикла;
	
	ВыборкаПоОР.Сбросить();
	
	Возврат ВыборкаПоОР;	
	
КонецФункции

Процедура ПроверитьЗаполненностьДанных(Выборка, СообщениеРезультатЗарузки)

	Если Выборка.ОбъектРемонтаКолВо > 1 Тогда				
		СообщениеРезультатЗарузки = СтрШаблон("%1В строке №%2 найдено более одного значения" +
		" для поля шаблона ""Объект ремонта"".%3", СообщениеРезультатЗарузки, Выборка.НомерСтроки, Символы.ПС);    
	КонецЕсли;  
	
	Если Выборка.ОбъектРемонта = Неопределено Тогда				
		СообщениеРезультатЗарузки = СтрШаблон("%1В строке №%2 не найдено значение" +
		" поля шаблона ""Объект ремонта"". Данные по строке не загружены.%3",
			СообщениеРезультатЗарузки, Выборка.НомерСтроки, Символы.ПС);			
	КонецЕсли; 
	
	Если Выборка.СтруктураИерархииКолВо > 1 Тогда				
		СообщениеРезультатЗарузки = СтрШаблон("%1В строке №%2 найдено более одного значения" +
		" для поля шаблона ""Структура иерархии"".%3", СообщениеРезультатЗарузки, Выборка.НомерСтроки, Символы.ПС);    
	КонецЕсли;  
	
	Если Выборка.СтруктураИерархии = Неопределено Тогда				
		СообщениеРезультатЗарузки = СтрШаблон("%1В строке №%2 не найдено значение" +
		" поля шаблона ""Структура иерархии"".%3",СообщениеРезультатЗарузки, Выборка.НомерСтроки, Символы.ПС); 
	КонецЕсли; 
	
	Если Выборка.ВидЭксплуатацииКолВо > 1 Тогда				
		СообщениеРезультатЗарузки = СтрШаблон("%1В строке №%2 найдено более одного значения" +
		" для поля шаблона ""Вид эксплуатации"".%3", СообщениеРезультатЗарузки, Выборка.НомерСтроки, Символы.ПС);    
	КонецЕсли;  
	
	Если Выборка.ВидЭксплуатации = Неопределено Тогда				
		СообщениеРезультатЗарузки = СтрШаблон("%1В строке №%2 не найдено значение" +
		" поля шаблона ""Вид эксплуатации"".%3",СообщениеРезультатЗарузки, Выборка.НомерСтроки, Символы.ПС); 
	КонецЕсли; 
	
	Если Выборка.ВидРемонтныхРаботКолВо > 1 Тогда				
		СообщениеРезультатЗарузки = СтрШаблон("%1В строке №%2 найдено более одного значения" +
		" для поля шаблона ""Вид ремонтных работ"".%3", СообщениеРезультатЗарузки, Выборка.НомерСтроки, Символы.ПС);    
	КонецЕсли;  
	
	Если Выборка.ВидРемонтныхРабот = Неопределено Тогда				
		СообщениеРезультатЗарузки = СтрШаблон("%1В строке №%2 не найдено значение" +
		" поля шаблона ""Вид ремонтных работа"".%3",СообщениеРезультатЗарузки, Выборка.НомерСтроки, Символы.ПС); 
	КонецЕсли; 
	
	Если Выборка.НаработкаПоказательКолВо > 1 Тогда				
		СообщениеРезультатЗарузки = СтрШаблон("%1В строке №%2 найдено более одного значения" +
		" для поля шаблона ""Наработка показатель"".%3", СообщениеРезультатЗарузки, Выборка.НомерСтроки, Символы.ПС);    
	КонецЕсли;  
	
	Если Выборка.НаработкаПоказатель = Неопределено Тогда				
		СообщениеРезультатЗарузки = СтрШаблон("%1В строке №%2 не найдено значение" +
		" поля шаблона ""Наработка показатель"".%3", СообщениеРезультатЗарузки, Выборка.НомерСтроки, Символы.ПС); 
	КонецЕсли; 
	
	СообщениеРезультатЗарузки = СообщениеРезультатЗарузки + Символы.ПС;	
	
КонецПроцедуры	

Функция СтрокаВДату(Знач Значение)   
	
	Значение = СтрЗаменить(Значение, " ", "");
	Значение = СокрЛП(СтрЗаменить(Значение, ".", ""));
	Значение = СокрЛП(СтрЗаменить(Значение, "/", "")); 
	Значение = СокрЛП(СтрЗаменить(Значение, "-", "")); 
	Значение = Сред(Значение, 5) + Сред(Значение, 3, 2) + Лев(Значение, 2);     
	
	Если СтрДлина(Значение) = 6 Тогда 
		Год = СтроковыеФункцииКлиентСервер.СтрокаВЧисло(Лев(Значение, 2));
		Если Год = Неопределено Тогда
			Возврат Дата('0001.01.01');
		КонецЕсли; 
		Значение = ?(Год > 29, "19", "20") + Значение; 
	КонецЕсли;     
	
	ОписаниеТипа = Новый ОписаниеТипов("Дата"); 
	Результат    = ОписаниеТипа.ПривестиЗначение(Значение);
	Возврат Результат;  
	
КонецФункции 

#КонецОбласти

//-- Проф-ИТ, #53, Башинская А., 13.10.2023	

#КонецЕсли