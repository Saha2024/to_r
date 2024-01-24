#Если Сервер Или ТолстыйКлиентОбычноеПриложение Или ВнешнееСоединение Тогда

#Область СлужебныеПроцедурыИФункции

// Заполняет список команд печати.
//
// Параметры:
// КомандыПечати – ТаблицаЗначений – состав полей см. в функции УправлениеПечатью.СоздатьКоллекциюКомандПечати.
//
Процедура ДобавитьКомандыПечати(КомандыПечати) Экспорт
	
	// Оценка работы подрядчика
	КомандаПечати = КомандыПечати.Добавить();
	КомандаПечати.МенеджерПечати = "Документ.торо_ОценкаРаботыПодрядчика";
	КомандаПечати.Идентификатор = "ОценкаКачестваРабот";
	КомандаПечати.Представление = НСтр("ru = 'Оценка работы подрядчика'");
	КомандаПечати.СразуНаПринтер = ОбщегоНазначенияВызовСервера.ХранилищеОбщихНастроекЗагрузить(
	"НастройкиТОиР",
	"ПечатьДокументовБезПредварительногоПросмотра",
	Ложь);
	
КонецПроцедуры

// Сформировать печатные формы объектов.
//
// ВХОДЯЩИЕ:
//   ИменаМакетов    - Строка    - Имена макетов, перечисленные через запятую.
//   МассивОбъектов  - Массив    - Массив ссылок на объекты которые нужно распечатать.
//   ПараметрыПечати - Структура - Структура дополнительных параметров печати.
//
// ИСХОДЯЩИЕ:
//   КоллекцияПечатныхФорм - Таблица значений - Сформированные табличные документы.
//   ПараметрыВывода       - Структура        - Параметры сформированных табличных документов.
//
Процедура Печать(МассивОбъектов, ПараметрыПечати, КоллекцияПечатныхФорм, ОбъектыПечати, ПараметрыВывода) Экспорт
	
	Если УправлениеПечатью.НужноПечататьМакет(КоллекцияПечатныхФорм, "ОценкаКачестваРабот") Тогда
		УправлениеПечатью.ВывестиТабличныйДокументВКоллекцию(КоллекцияПечатныхФорм, 
		"ОценкаКачестваРабот", 
		"Оценка качества работ подрядчика", 
		ПечатьОценкиКачестваРаботПодрядчика(МассивОбъектов, ПараметрыПечати),
		,
		"Документ.торо_ОценкаРаботыПодрядчика.ПФ_MXL_ОценкаКачестваРабот");
	КонецЕсли;
	
КонецПроцедуры

// Процедура вывода на экран печатной формы документа
// торо_ЗаявкаНаРемонт.
//
//  Параметры:
//   Заявка - ссылка на текущий документ торо_ЗаявкаНаРемонт.  
Функция ПечатьОценкиКачестваРаботПодрядчика(МассивОбъектов, ПараметрыПечати) 
	
	ТабДок = Новый ТабличныйДокумент;
	ТабДок.АвтоМасштаб = Истина;
		
	Макет = УправлениеПечатью.МакетПечатнойФормы("Документ.торо_ОценкаРаботыПодрядчика.ПФ_MXL_ОценкаКачестваРабот");

	ЭтоПерый = Истина;
	
	Для каждого Объект из МассивОбъектов Цикл
		
		Если НЕ ЭтоПерый Тогда
			ТабДок.ВывестиГоризонтальныйРазделительСтраниц();
		КонецЕсли;
		ЭтоПерый = Ложь;
		
		Запрос = Новый Запрос;
		Запрос.Текст = 
		"ВЫБРАТЬ РАЗРЕШЕННЫЕ
		|	торо_ОценкаРаботыПодрядчика.Ссылка КАК Ссылка,
		|	торо_ОценкаРаботыПодрядчика.Организация КАК Организация,
		|	торо_ОценкаРаботыПодрядчика.Контрагент КАК Контрагент,
		|	торо_ОценкаРаботыПодрядчика.АктОВыполненииРабот КАК АктОВыполненииРабот,
		|	торо_ОценкаРаботыПодрядчика.ОценкаДляВсехРабот КАК ОценкаДляВсехРабот
		|ПОМЕСТИТЬ ВТ_ДокОценка
		|ИЗ
		|	Документ.торо_ОценкаРаботыПодрядчика КАК торо_ОценкаРаботыПодрядчика
		|ГДЕ
		|	торо_ОценкаРаботыПодрядчика.Ссылка = &ДокОценка
		|
		|ИНДЕКСИРОВАТЬ ПО
		|	Ссылка
		|;
		|
		|////////////////////////////////////////////////////////////////////////////////
		|ВЫБРАТЬ РАЗРЕШЕННЫЕ РАЗЛИЧНЫЕ
		|	торо_АктОВыполненииЭтапаРаботРемонтыОборудования.ID КАК ID,
		|	торо_АктОВыполненииЭтапаРаботРемонтыОборудования.ОбъектРемонта КАК ОбъектРемонта,
		|	торо_АктОВыполненииЭтапаРаботРемонтыОборудования.ВидРемонтныхРабот КАК ВидРемонтныхРабот,
		|	торо_АктОВыполненииЭтапаРаботРемонтыОборудования.ДатаНачала КАК ДатаНачала,
		|	торо_АктОВыполненииЭтапаРаботРемонтыОборудования.ДатаОкончания КАК ДатаОкончания,
		|	торо_АктОВыполненииЭтапаРаботПодрядчикиРемонтныхРабот.Исполнитель КАК Контрагент
		|ПОМЕСТИТЬ ВТ_ДанныеАкт
		|ИЗ
		|	Документ.торо_АктОВыполненииЭтапаРабот.РемонтыОборудования КАК торо_АктОВыполненииЭтапаРаботРемонтыОборудования
		|		ЛЕВОЕ СОЕДИНЕНИЕ Документ.торо_АктОВыполненииЭтапаРабот.ИсполнителиПоРемонтам КАК торо_АктОВыполненииЭтапаРаботПодрядчикиРемонтныхРабот
		|		ПО торо_АктОВыполненииЭтапаРаботРемонтыОборудования.Ссылка = торо_АктОВыполненииЭтапаРаботПодрядчикиРемонтныхРабот.Ссылка
		|			И торо_АктОВыполненииЭтапаРаботРемонтыОборудования.ID = торо_АктОВыполненииЭтапаРаботПодрядчикиРемонтныхРабот.РемонтыОборудования_ID
		|			И (торо_АктОВыполненииЭтапаРаботПодрядчикиРемонтныхРабот.Исполнитель ССЫЛКА Справочник.Контрагенты)
		|		ВНУТРЕННЕЕ СОЕДИНЕНИЕ ВТ_ДокОценка КАК ВТ_ДокОценка
		|		ПО торо_АктОВыполненииЭтапаРаботРемонтыОборудования.Ссылка = ВТ_ДокОценка.АктОВыполненииРабот
		|ГДЕ
		|	торо_АктОВыполненииЭтапаРаботПодрядчикиРемонтныхРабот.Исполнитель В
		|			(ВЫБРАТЬ
		|				ВТ_ДокОценка.Контрагент КАК Контрагент
		|			ИЗ
		|				ВТ_ДокОценка КАК ВТ_ДокОценка)
		|
		|ИНДЕКСИРОВАТЬ ПО
		|	ID,
		|	ОбъектРемонта,
		|	ВидРемонтныхРабот
		|;
		|
		|////////////////////////////////////////////////////////////////////////////////
		|ВЫБРАТЬ РАЗРЕШЕННЫЕ
		|	ВТ_ДокОценка.Ссылка КАК Ссылка,
		|	ВТ_ДокОценка.Организация КАК Организация,
		|	ВТ_ДокОценка.Контрагент КАК Контрагент,
		|	ВТ_ДокОценка.АктОВыполненииРабот КАК АктОВыполненииРабот,
		|	торо_ОценкаРаботыПодрядчикаОбъектыРемонта.ID КАК ID,
		|	торо_ОценкаРаботыПодрядчикаОбъектыРемонта.ОбъектРемонта КАК ОбъектРемонта,
		|	торо_ОценкаРаботыПодрядчикаОбъектыРемонта.ВидРемонта КАК ВидРемонта
		|ПОМЕСТИТЬ Вт_ДанныеОценки2
		|ИЗ
		|	ВТ_ДокОценка КАК ВТ_ДокОценка
		|		ЛЕВОЕ СОЕДИНЕНИЕ Документ.торо_ОценкаРаботыПодрядчика.ОбъектыРемонта КАК торо_ОценкаРаботыПодрядчикаОбъектыРемонта
		|		ПО ВТ_ДокОценка.Ссылка = торо_ОценкаРаботыПодрядчикаОбъектыРемонта.Ссылка
		|ГДЕ
		|	ВТ_ДокОценка.ОценкаДляВсехРабот = ЛОЖЬ
		|
		|ОБЪЕДИНИТЬ ВСЕ
		|
		|ВЫБРАТЬ
		|	ВТ_ДокОценка.Ссылка,
		|	ВТ_ДокОценка.Организация,
		|	ВТ_ДокОценка.Контрагент,
		|	ВТ_ДокОценка.АктОВыполненииРабот,
		|	ВТ_ДанныеАкт.ID,
		|	ВТ_ДанныеАкт.ОбъектРемонта,
		|	ВТ_ДанныеАкт.ВидРемонтныхРабот
		|ИЗ
		|	ВТ_ДанныеАкт КАК ВТ_ДанныеАкт
		|		ВНУТРЕННЕЕ СОЕДИНЕНИЕ ВТ_ДокОценка КАК ВТ_ДокОценка
		|		ПО ВТ_ДанныеАкт.Контрагент = ВТ_ДокОценка.Контрагент
		|			И (ВТ_ДокОценка.ОценкаДляВсехРабот = ИСТИНА)
		|
		|ИНДЕКСИРОВАТЬ ПО
		|	Ссылка,
		|	ОбъектРемонта,
		|	ID
		|;
		|
		|////////////////////////////////////////////////////////////////////////////////
		|ВЫБРАТЬ РАЗРЕШЕННЫЕ
		|	Вт_ДанныеОценки.Ссылка КАК Ссылка,
		|	Вт_ДанныеОценки.Организация КАК Организация,
		|	Вт_ДанныеОценки.Контрагент КАК Контрагент,
		|	Вт_ДанныеОценки.ID КАК ID,
		|	Вт_ДанныеОценки.ОбъектРемонта КАК ОбъектРемонта,
		|	Вт_ДанныеОценки.ОбъектРемонта.Наименование КАК НаименованиеОР,
		|	Вт_ДанныеОценки.ВидРемонта КАК ВидРемонта,
		|	торо_ОценкаРаботыПодрядчикаПараметрыОценки.КритерийОценки КАК КритерийОценки,
		|	ВЫРАЗИТЬ(торо_ОценкаРаботыПодрядчикаПараметрыОценки.Примечание КАК СТРОКА(1000)) КАК Примечание,
		|	торо_ОценкаРаботыПодрядчикаПараметрыОценки.ЗначениеОценокПоКритериям КАК ЗначениеОценокПоКритериям
		|ПОМЕСТИТЬ Вт_ДанныеОценки
		|ИЗ
		|	Вт_ДанныеОценки2 КАК Вт_ДанныеОценки
		|		ЛЕВОЕ СОЕДИНЕНИЕ Документ.торо_ОценкаРаботыПодрядчика.ПараметрыОценки КАК торо_ОценкаРаботыПодрядчикаПараметрыОценки
		|		ПО Вт_ДанныеОценки.Ссылка = торо_ОценкаРаботыПодрядчикаПараметрыОценки.Ссылка
		|			И (Вт_ДанныеОценки.ID = торо_ОценкаРаботыПодрядчикаПараметрыОценки.ID
		|				ИЛИ торо_ОценкаРаботыПодрядчикаПараметрыОценки.ID = """")
		|
		|ИНДЕКСИРОВАТЬ ПО
		|	ID,
		|	ОбъектРемонта,
		|	ВидРемонта
		|;
		|
		|////////////////////////////////////////////////////////////////////////////////
		|ВЫБРАТЬ
		|	Вт_ДанныеОценки.Ссылка КАК Ссылка,
		|	Вт_ДанныеОценки.Организация КАК Организация,
		|	Вт_ДанныеОценки.Контрагент КАК Контрагент,
		|	Вт_ДанныеОценки.ID КАК ID,
		|	Вт_ДанныеОценки.ОбъектРемонта КАК ОбъектРемонта,
		|	Вт_ДанныеОценки.НаименованиеОР КАК НаименованиеОР,
		|	Вт_ДанныеОценки.ВидРемонта КАК ВидРемонта,
		|	ВТ_ДанныеАкт.ДатаНачала КАК ДатаНачала,
		|	ВТ_ДанныеАкт.ДатаОкончания КАК ДатаОкончания,
		|	Вт_ДанныеОценки.КритерийОценки КАК КритерийОценки,
		|	Вт_ДанныеОценки.ЗначениеОценокПоКритериям КАК ЗначениеОценокПоКритериям,
		|	Вт_ДанныеОценки.Примечание КАК Примечание
		|ПОМЕСТИТЬ ВТ_КритерииИДаты
		|ИЗ
		|	Вт_ДанныеОценки КАК Вт_ДанныеОценки
		|		ЛЕВОЕ СОЕДИНЕНИЕ ВТ_ДанныеАкт КАК ВТ_ДанныеАкт
		|		ПО Вт_ДанныеОценки.ID = ВТ_ДанныеАкт.ID
		|			И Вт_ДанныеОценки.ОбъектРемонта = ВТ_ДанныеАкт.ОбъектРемонта
		|			И Вт_ДанныеОценки.ВидРемонта = ВТ_ДанныеАкт.ВидРемонтныхРабот
		|
		|ИНДЕКСИРОВАТЬ ПО
		|	КритерийОценки
		|;
		|
		|////////////////////////////////////////////////////////////////////////////////
		|ВЫБРАТЬ
		|	ВТ_КритерииИДаты.Ссылка КАК Ссылка,
		|	ВТ_КритерииИДаты.Организация КАК Организация,
		|	ВТ_КритерииИДаты.Контрагент КАК Подрядчик,
		|	ВТ_КритерииИДаты.ОбъектРемонта КАК ОбъектРемонта,
		|	ВТ_КритерииИДаты.НаименованиеОР КАК НаименованиеОР,
		|	ВТ_КритерииИДаты.ВидРемонта КАК ВидРемонта,
		|	ВТ_КритерииИДаты.ДатаНачала КАК ДатаНачала,
		|	ВТ_КритерииИДаты.ДатаОкончания КАК ДатаОкончания,
		|	ВТ_КритерииИДаты.КритерийОценки КАК КритерийОценки,
		|	ВТ_КритерииИДаты.ЗначениеОценокПоКритериям КАК Оценка,
		|	торо_КритерииОценкиБазовыеКритерииВыбораПодрядчика.КритерийВыбораПодрядчика КАК КритерийВыбораПодрядчика,
		|	торо_КритерииОценкиБазовыеКритерииВыбораПодрядчика.КритерийВыбораПодрядчика.Комментарий КАК Комментарий,
		|	ВТ_КритерииИДаты.Примечание КАК Примечание
		|ИЗ
		|	ВТ_КритерииИДаты КАК ВТ_КритерииИДаты
		|		ЛЕВОЕ СОЕДИНЕНИЕ Справочник.торо_КритерииОценки.БазовыеКритерииВыбораПодрядчика КАК торо_КритерииОценкиБазовыеКритерииВыбораПодрядчика
		|		ПО ВТ_КритерииИДаты.КритерийОценки = торо_КритерииОценкиБазовыеКритерииВыбораПодрядчика.Ссылка
		|ИТОГИ
		|	МАКСИМУМ(Подрядчик),
		|	МАКСИМУМ(НаименованиеОР),
		|	МАКСИМУМ(ДатаНачала),
		|	МАКСИМУМ(ДатаОкончания),
		|	МАКСИМУМ(Оценка),
		|	МАКСИМУМ(Примечание)
		|ПО
		|	ОбъектРемонта,
		|	ВидРемонта,
		|	КритерийОценки";
		
		Запрос.УстановитьПараметр("ДокОценка", Объект);
		
		обЗаголовок = Макет.ПолучитьОбласть("Заголовок");
		обШапка = Макет.ПолучитьОбласть("Шапка");
		обШапкаТаблицы = Макет.ПолучитьОбласть("ШапкаТаблицы");
		обСтрокаТаблицы = Макет.ПолучитьОбласть("СтрокаТаблицы");
		обПодвал = Макет.ПолучитьОбласть("Подвал");
		обПустаяСтрока = Макет.ПолучитьОбласть("ПустаяСтрока");
		
				
		обЗаголовок.Параметры.Организация = торо_ЗаполнениеДокументов.ПолучитьПредставлениеОрганизацииДляПечати(Объект.Организация);;
		обЗаголовок.Параметры.Дата = Объект.Дата;
		обЗаголовок.Параметры.НомерДок = Объект.Номер;
		ТабДок.Вывести(обЗаголовок);
		
		Результат = Запрос.Выполнить();
		Если Результат.Пустой() Тогда
			ТабДок.Вывести(обПодвал);
			Продолжить;
		КонецЕсли;
		
		ВыборкаОР = Результат.Выбрать(ОбходРезультатаЗапроса.ПоГруппировкам);
		Пока ВыборкаОР.Следующий() цикл
			ВыборкаВР = ВыборкаОР.Выбрать(ОбходРезультатаЗапроса.ПоГруппировкам);
			Пока ВыборкаВР.Следующий() Цикл
				
				обШапка = Макет.ПолучитьОбласть("Шапка");
				ЗаполнитьЗначенияСвойств(обШапка.Параметры, ВыборкаВР);
				обШапка.Параметры.ДатаНачала = Формат(ВыборкаВР.ДатаНачала, "ДФ=dd.MM.yyyy");
				обШапка.Параметры.ДатаОкончания = Формат(ВыборкаВР.ДатаОкончания, "ДФ=dd.MM.yyyy");
				обШапка.Параметры.НаименованиеРаботы = "" + ВыборкаВР.НаименованиеОР + " - " + ВыборкаВР.ВидРемонта;
				ТабДок.Вывести(обШапка);
				
				ТабДок.Вывести(обШапкаТаблицы);
				
				ДетЗаписи = ВыборкаВР.Выбрать(ОбходРезультатаЗапроса.ПоГруппировкам);
				НомерПП = 1;
				Пока ДетЗаписи.Следующий() Цикл
					
					обСтрокаТаблицы = Макет.ПолучитьОбласть("СтрокаТаблицы");
					ЗаполнитьЗначенияСвойств(обСтрокаТаблицы.Параметры, ДетЗаписи);
					
					СписокКритериев = ДетЗаписи.Выбрать();
					
					КритерийПодрядчика = "";
					Пока СписокКритериев.Следующий() Цикл
						КритерийПодрядчика = КритерийПодрядчика
						+ СписокКритериев.КритерийВыбораПодрядчика + " - " 
						+ СписокКритериев.Комментарий + "; " + Символы.ПС;
					КонецЦикла;
					обСтрокаТаблицы.Параметры.КритерийПодрядчика = КритерийПодрядчика;
					обСтрокаТаблицы.Параметры.НомерПП = НомерПП;
					
					табДок.Вывести(обСтрокаТаблицы);
					
					НомерПП = НомерПП + 1;
					
				КонецЦикла;
				ТабДок.Вывести(обПустаяСтрока);
			КонецЦикла;
		КонецЦикла;
		
		обПодвал.Параметры.Организация = Объект.Организация;
		ТабДок.Вывести(обПодвал);
		
	КонецЦикла;
	
	ТабДок.ТолькоПросмотр = Истина;
	ТабДок.КлючПараметровПечати = "торо_ПечатьОценкиКачестваРаботПодрядчика";
	Возврат ТабДок;

КонецФункции

Процедура ПриОпределенииНастроекВерсионированияОбъектов(Настройки) Экспорт
	
КонецПроцедуры

// СтандартныеПодсистемы.УправлениеДоступом

// См. УправлениеДоступомПереопределяемый.ПриЗаполненииСписковСОграничениемДоступа.
Процедура ПриЗаполненииОграниченияДоступа(Ограничение) Экспорт
	
	Ограничение.Текст =
	"РазрешитьЧтение
	|ГДЕ 
	|	ЗначениеРазрешено(Организация)
	|	И ЗначениеРазрешено(Подразделение)
	|;
	|РазрешитьИзменениеЕслиРазрешеноЧтение
	|ГДЕ
	|	ДляВсехСтрок(&Ограничение_ОР)";

	ОграничениеОР = торо_УправлениеДоступом.ПолучитьОграничениеДоступаДляОбъектаРемонтаВТаблице("ОбъектыРемонта.ОбъектРемонта", Истина);
	Ограничение.Текст = СтрЗаменить(Ограничение.Текст, "&Ограничение_ОР", ОграничениеОР);

	Ограничение.ТекстДляВнешнихПользователей = Ограничение.Текст;

КонецПроцедуры

// Конец СтандартныеПодсистемы.УправлениеДоступом

#КонецОбласти

#КонецЕсли