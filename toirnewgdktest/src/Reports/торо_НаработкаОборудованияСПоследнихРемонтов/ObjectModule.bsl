#Если Сервер Или ТолстыйКлиентОбычноеПриложение Или ВнешнееСоединение Тогда

#Область ОбработчикиСобытий

Процедура ПриКомпоновкеРезультата(ДокументРезультат, ДанныеРасшифровки, СтандартнаяОбработка)
	
	СтандартнаяОбработка = Ложь;
	
	ТекущаяИерархия = торо_ОтчетыСервер.ПолучитьЗначениеСтруктурыИерархии(КомпоновщикНастроек);
	
	торо_ОтчетыСервер.УстановитьЗапросыНаборовДанныхИерархииОР(СхемаКомпоновкиДанных, ТекущаяИерархия, "ТекущаяДата");

	Если ТекущаяИерархия.СтроитсяАвтоматически Тогда
		торо_ОтчетыКлиентСервер.УстановитьТипИерархическойГруппировкиВНастройках(КомпоновщикНастроек, "ОбъектИерархии", ТипГруппировкиКомпоновкиДанных.Иерархия);
	Иначе
		торо_ОтчетыКлиентСервер.УстановитьТипИерархическойГруппировкиВНастройках(КомпоновщикНастроек, "ОбъектИерархии", ТипГруппировкиКомпоновкиДанных.ТолькоИерархия);
	КонецЕсли;
		
	// Сгенерируем макет компоновки данных при помощи компоновщика макета.
	КомпоновщикМакета = Новый КомпоновщикМакетаКомпоновкиДанных;
	Настройки = КомпоновщикНастроек.ПолучитьНастройки();
	
	ТабНаработки = ПолучитьТабЗначСРезультатомОтчета(ТекущаяИерархия);
	
	ФОВыделятьИнвестиционныеРемонты = ПолучитьФункциональнуюОпцию("торо_ВыделятьИнвестиционныеРемонты");
	Для Каждого ЭлементВыбора Из Настройки.Выбор.Элементы Цикл
		Если ЭлементВыбора.Заголовок = "Последний капитальный ремонт" Тогда
			ЭлементВыбора.Использование = ФОВыделятьИнвестиционныеРемонты;
		КонецЕсли;
	КонецЦикла;
		
	// В качестве схемы компоновки будет выступать схема самого отчета.
	// В качестве настроек отчета - текущие настройки отчета.
	// Данные расшифровки будем помещать в свойство расширения формы - ДанныеРасшифровки.
	МакетКомпоновки = КомпоновщикМакета.Выполнить(СхемаКомпоновкиДанных, Настройки, ДанныеРасшифровки);
	
	// получаем внешние наборы данных
	ВнешниеНаборыДанных = Новый Структура;
	ВнешниеНаборыДанных.Вставить("НаборДанных", ТабНаработки);
	
	// Создадим и инициализируем процессор компоновки
	ПроцессорКомпоновки = Новый ПроцессорКомпоновкиДанных;
	ПроцессорКомпоновки.Инициализировать(МакетКомпоновки, ВнешниеНаборыДанных,ДанныеРасшифровки,Истина);
	
	// Создадим и инициализируем процессор вывода результата
	ПроцессорВывода = Новый ПроцессорВыводаРезультатаКомпоновкиДанныхВТабличныйДокумент;
	ПроцессорВывода.УстановитьДокумент(ДокументРезультат);
	ПроцессорВывода.ОтображатьПроцентВывода = Истина;

	// выводим отчет
	ПроцессорВывода.Вывести(ПроцессорКомпоновки,Истина);
	
КонецПроцедуры

#КонецОбласти

#Область СлужебныеПроцедурыИФункции

// Настройки общей формы отчета подсистемы "Варианты отчетов".
//
// Параметры:
//   Форма - ФормаКлиентскогоПриложения - Форма отчета.
//   КлючВарианта - Строка - Имя предопределенного варианта отчета или уникальный идентификатор пользовательского.
//   Настройки - Структура - см. возвращаемое значение ФункцииОтчетовКлиентСервер.ПолучитьНастройкиОтчетаПоУмолчанию().
//
Процедура ОпределитьНастройкиФормы(Форма, КлючВарианта, Настройки) Экспорт
	Настройки.События.ПриЗагрузкеПользовательскихНастроекНаСервере = Истина;
	Настройки.События.ПриЗагрузкеВариантаНаСервере = Истина;
	
	Настройки.ЗагрузитьНастройкиПриИзмененииПараметров = ЗагрузитьНастройкиПриИзмененииПараметров();
КонецПроцедуры

Процедура ПриЗагрузкеВариантаНаСервере(Форма, НовыеПользовательскиеНастройкиКД) Экспорт
	ПользовательскиеНастройкиЭлементы = КомпоновщикНастроек.ПользовательскиеНастройки.Элементы;
	ВыбратьЗначенияНастроек(ПользовательскиеНастройкиЭлементы, Форма);
КонецПроцедуры

Процедура ПриЗагрузкеПользовательскихНастроекНаСервере(Форма, НовыеПользовательскиеНастройкиКД) Экспорт
	ПользовательскиеНастройкиЭлементы = КомпоновщикНастроек.ПользовательскиеНастройки.Элементы;
	ВыбратьЗначенияНастроек(ПользовательскиеНастройкиЭлементы, Форма);
КонецПроцедуры

Процедура ВыбратьЗначенияНастроек(ПользовательскиеНастройкиЭлементы, Форма)
	
	ПараметрыВыполнения = Форма.Параметры;
	
	Для каждого Элемент Из ПользовательскиеНастройкиЭлементы Цикл
		Если ТипЗнч(Элемент) = Тип("ЗначениеПараметраНастроекКомпоновкиДанных") Тогда
			Если Строка(Элемент.Параметр) = "ОбъектРемонта" Тогда
				Если ПараметрыВыполнения.Свойство("ОбъектРемонта") Тогда
					Элемент.Значение      = ПараметрыВыполнения.ОбъектРемонта;
					Элемент.Использование = Истина;			
					Продолжить;
				КонецЕсли;
			ИначеЕсли Строка(Элемент.Параметр) = "СтруктураИерархии" Тогда
				Если ПараметрыВыполнения.Свойство("СтруктураИерархии") Тогда
					Элемент.Значение      = ПараметрыВыполнения.СтруктураИерархии;
					Элемент.Использование = Истина;			
					Продолжить;
				КонецЕсли;
			ИначеЕсли Строка(Элемент.Параметр) = "ВместеСПодчиненными" Тогда
				Если ПараметрыВыполнения.Свойство("ВместеСПодчиненными") Тогда
					Элемент.Значение      = ПараметрыВыполнения.ВместеСПодчиненными;
					Элемент.Использование = Истина;			
					Продолжить;
				КонецЕсли;						
			КонецЕсли;
		КонецЕсли;
	КонецЦикла;
КонецПроцедуры 

Функция ЗагрузитьНастройкиПриИзмененииПараметров()  
	
	Параметры = Новый Массив;
	Параметры.Добавить(Новый ПараметрКомпоновкиДанных("ИерархияТип"));	
	Возврат Параметры;
	
КонецФункции

Функция ПолучитьТабЗначСРезультатомОтчета(ТекущаяИерархия)
	
	Период = ПолучитьПользовательскуюНастройкуКомпоновщика(КомпоновщикНастроек, "Период");
	ДатаНач = Период.Значение.ДатаНачала;
	
	Если Период.Использование = Истина Тогда // Проверка на использование параметра
		ДатаКон = Период.Значение.ДатаОкончания;
	Иначе
		ДатаКон = Неопределено;
	КонецЕсли;
	
	РемонтыПоОтбору = ПолучитьРемонтыПоПользовательскомуОтбору();
	
	Запрос = Новый Запрос;
	МенВТ = Новый МенеджерВременныхТаблиц;
	Запрос.МенеджерВременныхТаблиц = МенВТ;
	Запрос.Текст = "ВЫБРАТЬ
	|	ПоследниеРемонты.ОбъектРемонтныхРабот КАК ОбъектРемонта,
	|	ПоследниеРемонты.ВидРемонтныхРабот,
	|	ПоследниеРемонты.ID,                                                                                                                                    
	|	ПоследниеРемонты.ДатаКон КАК ДатаОкончания
	|ПОМЕСТИТЬ ПоследниеРемонты
	|ИЗ
	|	&ПоследниеРемонты КАК ПоследниеРемонты
	|;
	|
	|////////////////////////////////////////////////////////////////////////////////
	|ВЫБРАТЬ
	|	ПоследниеРемонты.ОбъектРемонта,
	|	МАКСИМУМ(ПоследниеРемонты.ДатаОкончания) КАК ДатаОкончания
	|ПОМЕСТИТЬ ТабМаксДат
	|ИЗ
	|	ПоследниеРемонты КАК ПоследниеРемонты
	|
	|СГРУППИРОВАТЬ ПО
	|	ПоследниеРемонты.ОбъектРемонта
	|;
	|						
	|////////////////////////////////////////////////////////////////////////////////
	|ВЫБРАТЬ
	|	ПоследниеРемонты.ОбъектРемонта,
	|	ПоследниеРемонты.ВидРемонтныхРабот КАК ВидРемонта,
	|	ПоследниеРемонты.ID КАК ID_Ремонта,
	|	ПоследниеРемонты.ДатаОкончания КАК ДатаНачала
	|ПОМЕСТИТЬ Капитальные
	|ИЗ
	|	ПоследниеРемонты КАК ПоследниеРемонты
	|ГДЕ
	|	ПоследниеРемонты.ВидРемонтныхРабот В
	|			(ВЫБРАТЬ
	|				торо_ВидыРемонтов.Ссылка
	|			ИЗ
	|				Справочник.торо_ВидыРемонтов КАК торо_ВидыРемонтов
	|			ГДЕ
	|				торо_ВидыРемонтов.КапитальныйРемонт)
	|;
	|
	|////////////////////////////////////////////////////////////////////////////////
	|ВЫБРАТЬ
	|	ПоследниеРемонты.ОбъектРемонта,
	|	ПоследниеРемонты.ВидРемонтныхРабот,
	|	ПоследниеРемонты.ID,
	|	ПоследниеРемонты.ДатаОкончания
	|ПОМЕСТИТЬ НеКапитальные
	|ИЗ
	|	ТабМаксДат КАК ТабМаксДат
	|		ВНУТРЕННЕЕ СОЕДИНЕНИЕ ПоследниеРемонты КАК ПоследниеРемонты
	|		ПО ТабМаксДат.ОбъектРемонта = ПоследниеРемонты.ОбъектРемонта
	|			И ТабМаксДат.ДатаОкончания = ПоследниеРемонты.ДатаОкончания
	|;
	|
	|////////////////////////////////////////////////////////////////////////////////
	|ВЫБРАТЬ
	|	ЕСТЬNULL(Капитальные.ОбъектРемонта, НеКапитальные.ОбъектРемонта) КАК ОбъектРемонта,
	|	Капитальные.ВидРемонта КАК ВидРемонтаКапитальный,
	|	НеКапитальные.ВидРемонтныхРабот,
	|	Капитальные.ДатаНачала КАК ДатаРемонтаКапитальный,
	|	НеКапитальные.ДатаОкончания КАК ДатаРемонта,
	|	Капитальные.ID_Ремонта КАК IDКапитального,
	|	НеКапитальные.ID
	|ПОМЕСТИТЬ Результат
	|ИЗ
	|	Капитальные КАК Капитальные
	|		ПОЛНОЕ СОЕДИНЕНИЕ НеКапитальные КАК НеКапитальные
	|		ПО Капитальные.ОбъектРемонта = НеКапитальные.ОбъектРемонта
	|;
	|
	|////////////////////////////////////////////////////////////////////////////////
	|ВЫБРАТЬ
	|	Капитальные.ОбъектРемонта,
	|	Капитальные.ВидРемонта,
	|	Капитальные.ID_Ремонта,
	|	Капитальные.ДатаНачала
	|ИЗ
	|	Капитальные КАК Капитальные
	|
	|ОБЪЕДИНИТЬ
	|
	|ВЫБРАТЬ
	|	НеКапитальные.ОбъектРемонта,
	|	НеКапитальные.ВидРемонтныхРабот,
	|	НеКапитальные.ID,
	|	НеКапитальные.ДатаОкончания
	|ИЗ
	|	НеКапитальные КАК НеКапитальные";
	
	Запрос.УстановитьПараметр("ПоследниеРемонты", РемонтыПоОтбору);
	
	ТабРемонтов = Запрос.Выполнить().Выгрузить();
	
	ТабНаработки = торо_РаботаСНаработкой.РассчитатьНаработкуСРемонта(ТабРемонтов, ДатаНач, ДатаКон);
	
	Запрос = Новый Запрос;
	Запрос.МенеджерВременныхТаблиц = МенВТ;
	Запрос.Текст = 
		"ВЫБРАТЬ
		|	ТабНаработки.ДатаНач,
		|	ТабНаработки.ОбъектРемонтныхРабот,
		|	ТабНаработки.ВидРемонтныхРабот,
		|	ТабНаработки.ID,
		|	ТабНаработки.Наработка,
		|	ТабНаработки.Показатель
		|ПОМЕСТИТЬ ТабНаработки
		|ИЗ
		|	&ТабНаработки КАК ТабНаработки
		|;
		|
		|////////////////////////////////////////////////////////////////////////////////
		|ВЫБРАТЬ
		|	Результат.ДатаРемонта КАК ДатаКон,
		|	Результат.ДатаРемонтаКапитальный КАК ДатаКонКапитальный,
		|	Результат.ОбъектРемонта КАК ОбъектРемонтныхРабот,
		|	Результат.ВидРемонтныхРабот,
		|	Результат.ВидРемонтаКапитальный,
		|	Результат.IDКапитального,
		|
		|	ТабНаработки.Наработка КАК Наработка,
		|	ТабНаработки.Показатель КАК Показатель
		|ПОМЕСТИТЬ ТабСНаработкой
		|ИЗ
		|	Результат КАК Результат
		|		ЛЕВОЕ СОЕДИНЕНИЕ ТабНаработки КАК ТабНаработки
		|		ПО Результат.ID = ТабНаработки.ID
		|
		|;
		|////////////////////////////////////////////////////////////////////////////////
		|ВЫБРАТЬ
		|	ТабСНаработкой.ДатаКон,
		|	ТабСНаработкой.ДатаКонКапитальный,
		|	ТабСНаработкой.ОбъектРемонтныхРабот.Организация КАК Организация,
		|	ТабСНаработкой.ОбъектРемонтныхРабот.Подразделение КАК Подразделение,
		|	ТабСНаработкой.ОбъектРемонтныхРабот,
		|	ТабСНаработкой.ВидРемонтныхРабот,
		|	ТабСНаработкой.ВидРемонтаКапитальный,
		|	ТабСНаработкой.Наработка,
		|	ТабСНаработкой.Показатель,
		|	ТабНаработки.Наработка КАК НаработкаКапитальный,
		|	ТабНаработки.Показатель КАК ПоказательКапитальный,
		|	&ОбъектИерархии КАК ОбъектИерархии
		|ИЗ
		|	ТабСНаработкой КАК ТабСНаработкой
		|		ЛЕВОЕ СОЕДИНЕНИЕ ТабНаработки КАК ТабНаработки
		|		ПО ТабСНаработкой.IDКапитального = ТабНаработки.ID
		|		И ТабСНаработкой.Показатель = ТабНаработки.Показатель";
		
	Если ТекущаяИерархия.СтроитсяАвтоматически Тогда
		ПараметрОбъектИерархии = СтрШаблон("ВЫРАЗИТЬ(ТабСНаработкой.ОбъектРемонтныхРабот КАК Справочник.торо_ОбъектыРемонта).%1", ТекущаяИерархия.РеквизитОР);              
	Иначе
		ПараметрОбъектИерархии = "ТабСНаработкой.ОбъектРемонтныхРабот";
	КонецЕсли;
	
	Запрос.Текст = СтрЗаменить(Запрос.Текст, "&ОбъектИерархии", ПараметрОбъектИерархии);
	
	Запрос.УстановитьПараметр("ТабНаработки", ТабНаработки);
	РезультатЗапроса = Запрос.Выполнить().Выгрузить();
	
	Возврат РезультатЗапроса;
	
КонецФункции

Функция НайтиДанные(ИмяПоиска, МестоПоиска)
	
	Для каждого Элемент Из МестоПоиска Цикл
		Если Элемент.Имя = ИмяПоиска Тогда
			
			Возврат Элемент;
			
		Иначе
			
			Результат = НайтиДанные(ИмяПоиска, Элемент.Структура);
			Если Результат <> Неопределено Тогда
				Возврат Результат;
			КонецЕсли;
			
		КонецЕсли;
		
	КонецЦикла;
	
КонецФункции

Функция ПолучитьРемонтыПоПользовательскомуОтбору()
	
	ТекущаяИерархия = торо_ОтчетыСервер.ПолучитьЗначениеСтруктурыИерархии(КомпоновщикНастроек);
	
	СхемаКомпоновкиОтбор = ПолучитьМакет("СхемаКомпоновкиОтбор");
	
	ПараметрыИсходнойСхемы = СхемаКомпоновкиДанных.Параметры;
	Для каждого ПараметрИсходнойСхемы из ПараметрыИсходнойСхемы Цикл
		
		ПараметрСхемыОтбора = СхемаКомпоновкиОтбор.Параметры.Найти(ПараметрИсходнойСхемы.Имя);
		Если ПараметрСхемыОтбора = Неопределено Тогда
			ПараметрСхемыОтбора = СхемаКомпоновкиОтбор.Параметры.Добавить();
			ПараметрСхемыОтбора.Имя = ПараметрИсходнойСхемы.Имя;
		КонецЕсли;
		
		ЗаполнитьЗначенияСвойств(ПараметрСхемыОтбора, ПараметрыИсходнойСхемы);
	КонецЦикла;
	
	КомпоновщикСхемыОтбора = Новый КомпоновщикНастроекКомпоновкиДанных;
	КомпоновщикСхемыОтбора.Инициализировать(Новый ИсточникДоступныхНастроекКомпоновкиДанных(СхемаКомпоновкиОтбор));
	КомпоновщикСхемыОтбора.ЗагрузитьНастройки(СхемаКомпоновкиОтбор.НастройкиПоУмолчанию);
	
	ИспользованиеОтбораПоОбъектуРемонта = ПолучитьПользовательскуюНастройкуКомпоновщика(КомпоновщикНастроек, "ОбъектРемонта", "Использование");
	Если ИспользованиеОтбораПоОбъектуРемонта Тогда
		
		ОбъектРемонта = ПолучитьПользовательскуюНастройкуКомпоновщика(КомпоновщикНастроек, "ОбъектРемонта", "Значение");
		ВместеСПодчиненными = ПолучитьПользовательскуюНастройкуКомпоновщика(КомпоновщикНастроек, "ВместеСПодчиненными", "Значение"); 
		
		Если ВместеСПодчиненными Тогда
			МассивОбъектовРемонта = торо_РаботаСИерархией20.ПолучитьМассивПодчиненныхОбъектов(ОбъектРемонта, ТекущаяИерархия); 
		Иначе
			МассивОбъектовРемонта = Новый Массив();
		КонецЕсли;
		МассивОбъектовРемонта.Добавить(ОбъектРемонта);
		
		ПараметрОбъектыРемонта = КомпоновщикСхемыОтбора.Настройки.ПараметрыДанных.НайтиЗначениеПараметра(Новый ПараметрКомпоновкиДанных("ОбъектыРемонта"));
		
		ПараметрОбъектыРемонта.Значение.ЗагрузитьЗначения(МассивОбъектовРемонта);	
		ПараметрОбъектыРемонта.Использование = Истина;
	КонецЕсли;	
			
	
	ТекстЗапроса = "ВЫБРАТЬ РАЗРЕШЕННЫЕ  
	|	торо_ЗавершенныеРемонтныеРаботыСрезПоследних.ОбъектРемонта.Организация КАК Организация,
	|	торо_ЗавершенныеРемонтныеРаботыСрезПоследних.ОбъектРемонта.Подразделение КАК Подразделение,
	|	торо_ЗавершенныеРемонтныеРаботыСрезПоследних.ОбъектРемонта КАК ОбъектРемонтныхРабот,
	|	торо_ЗавершенныеРемонтныеРаботыСрезПоследних.ВидРемонтныхРабот КАК ВидРемонтныхРабот,
	|	торо_ЗавершенныеРемонтныеРаботыСрезПоследних.ID КАК ID,
	|	торо_ЗавершенныеРемонтныеРаботыСрезПоследних.ДатаОкончания КАК ДатаКон
	|ПОМЕСТИТЬ Ремонты
	|ИЗ
	|	РегистрСведений.торо_ЗавершенныеРемонтныеРаботы.СрезПоследних({(&ДатаОкончания)}, {(ОбъектРемонта В (&ОбъектыРемонта))}) КАК торо_ЗавершенныеРемонтныеРаботыСрезПоследних
	|;
	|
	|////////////////////////////////////////////////////////////////////////////////
	|ВЫБРАТЬ РАЗРЕШЕННЫЕ
	|	Ремонты.Организация КАК Организация,
	|	Ремонты.Подразделение КАК Подразделение,
	|	Ремонты.ОбъектРемонтныхРабот КАК ОбъектРемонтныхРабот,
	|	Ремонты.ВидРемонтныхРабот КАК ВидРемонтныхРабот,
	|	Ремонты.ID КАК ID,
	|	Ремонты.ДатаКон КАК ДатаКон,
	|	&ОбъектИерархии КАК ОбъектИерархии
	|ИЗ
	|	Ремонты КАК Ремонты
	|		ВНУТРЕННЕЕ СОЕДИНЕНИЕ Справочник.торо_ОбъектыРемонта КАК торо_ОбъектыРемонта
	|		ПО Ремонты.ОбъектРемонтныхРабот = торо_ОбъектыРемонта.Ссылка";
	
	Если ТекущаяИерархия.СтроитсяАвтоматически Тогда
		ПараметрОбъектИерархии = СтрШаблон("ВЫРАЗИТЬ(Ремонты.ОбъектРемонтныхРабот КАК Справочник.торо_ОбъектыРемонта).%1", ТекущаяИерархия.РеквизитОР);              
	Иначе
		ПараметрОбъектИерархии = "Ремонты.ОбъектРемонтныхРабот";
	КонецЕсли;
	
	ТекстЗапроса = СтрЗаменить(ТекстЗапроса, "&ОбъектИерархии", ПараметрОбъектИерархии);
		
	СхемаКомпоновкиОтбор.НаборыДанных.Объекты.Запрос = ТекстЗапроса;
	СхемаКомпоновкиОтбор.НаборыДанных.Иерархия.Запрос = СхемаКомпоновкиДанных.НаборыДанных.Иерархия.Запрос;
	СхемаКомпоновкиОтбор.НаборыДанных.Контроль.Запрос = СхемаКомпоновкиДанных.НаборыДанных.Контроль.Запрос;
		
	НастройкиИсходнойСхемы = КомпоновщикНастроек.ПолучитьНастройки(); 	
	Для каждого ПараметрИсходнойСхемы из НастройкиИсходнойСхемы.ПараметрыДанных.Элементы Цикл
		
		ПараметрСхемыОтбора = КомпоновщикСхемыОтбора.Настройки.ПараметрыДанных.НайтиЗначениеПараметра(ПараметрИсходнойСхемы.Параметр);
		Если ПараметрСхемыОтбора <> Неопределено Тогда
			КомпоновщикСхемыОтбора.Настройки.ПараметрыДанных.УстановитьЗначениеПараметра(ПараметрИсходнойСхемы.Параметр, ПараметрИсходнойСхемы.Значение);	
			ПараметрСхемыОтбора.Использование = ПараметрИсходнойСхемы.Использование;
		КонецЕсли;
	КонецЦикла;
			
	Для каждого ЭлементОтбораИсходнойСхемы из НастройкиИсходнойСхемы.Отбор.Элементы Цикл
		
		ТипОтбораЭлемент = ТипЗнч(ЭлементОтбораИсходнойСхемы) = Тип("ЭлементОтбораКомпоновкиДанных");
		
		Если ТипОтбораЭлемент Тогда
			ПолеНеДоступноДляОтбора = КомпоновщикСхемыОтбора.Настройки.ДоступныеПоляОтбора.НайтиПоле(ЭлементОтбораИсходнойСхемы.ЛевоеЗначение) = Неопределено;
			Если ПолеНеДоступноДляОтбора Тогда
				Продолжить;
			КонецЕсли; 
		КонецЕсли;	
		
		НовыйОтбор = КомпоновщикСхемыОтбора.Настройки.Отбор.Элементы.Добавить(ТипЗнч(ЭлементОтбораИсходнойСхемы));
		ЗаполнитьЗначенияСвойств(НовыйОтбор, ЭлементОтбораИсходнойСхемы);
		Если ТипОтбораЭлемент Тогда
			Если ТипЗнч(ЭлементОтбораИсходнойСхемы.ПравоеЗначение) = Тип("СписокЗначений") Тогда
				НовыйОтбор.ПравоеЗначение.ЗагрузитьЗначения(ЭлементОтбораИсходнойСхемы.ПравоеЗначение.ВыгрузитьЗначения());
			КонецЕсли;
		КонецЕсли;
	КонецЦикла;
		
	Настройки = КомпоновщикСхемыОтбора.ПолучитьНастройки();
		
	КомпоновщикМакета = Новый КомпоновщикМакетаКомпоновкиДанных;
	МакетКомпоновки = КомпоновщикМакета.Выполнить(СхемаКомпоновкиОтбор, Настройки,,,Тип("ГенераторМакетаКомпоновкиДанныхДляКоллекцииЗначений"));
	
	ПроцессорКомпоновки = Новый ПроцессорКомпоновкиДанных;
	ПроцессорКомпоновки.Инициализировать(МакетКомпоновки);
	
	РемонтыСОтбором = Новый ТаблицаЗначений;
	
	ПроцессорВывода = Новый ПроцессорВыводаРезультатаКомпоновкиДанныхВКоллекциюЗначений;
	ПроцессорВывода.УстановитьОбъект(РемонтыСОтбором);
	ПроцессорВывода.Вывести(ПроцессорКомпоновки);
	
	Возврат РемонтыСОтбором;	
	
КонецФункции

Функция ПолучитьПользовательскуюНастройкуКомпоновщика(Компоновщик, ИмяНастройки, СвойствоНастройки = Неопределено)

	ЗначениеПараметраКомпоновкиДанных = Компоновщик.Настройки.ПараметрыДанных.НайтиЗначениеПараметра(Новый ПараметрКомпоновкиДанных(ИмяНастройки)); 
	Если ЗначениеПараметраКомпоновкиДанных = Неопределено Тогда
		Возврат Неопределено;
	КонецЕсли;	
	
	ИдентификаторНастройки = Компоновщик.ПользовательскиеНастройки.ПолучитьИдентификаторПоОбъекту(ЗначениеПараметраКомпоновкиДанных);
    Если ИдентификаторНастройки = Неопределено Тогда
		Возврат Неопределено;
	КонецЕсли;
	
	Если СвойствоНастройки = Неопределено Тогда
		Возврат Компоновщик.ПользовательскиеНастройки.Элементы.Найти(ИдентификаторНастройки);
	Иначе
		Возврат Компоновщик.ПользовательскиеНастройки.Элементы.Найти(ИдентификаторНастройки)[СвойствоНастройки];
	КонецЕсли;
	
КонецФункции

#КонецОбласти

#КонецЕсли