#Область ОбработчикиСобытий

&ИзменениеИКонтроль("ОбработкаПроведения")
Процедура проф_ОбработкаПроведения(Отказ, РежимПроведения)

	МассивДокументовОснований = ОбщегоНазначения.ВыгрузитьКолонку(ДокументыОснования, "ДокументОснование");

	Запрос = Новый Запрос( "ВЫБРАТЬ РАЗРЕШЕННЫЕ
	|	торо_ВнешнееОснованиеДляРабот.Ссылка КАК Ссылка,
	|	торо_ВнешнееОснованиеДляРабот.Проведен КАК Проведен
	|ИЗ
	|	Документ.торо_ВнешнееОснованиеДляРабот КАК торо_ВнешнееОснованиеДляРабот
	|ГДЕ
	|	торо_ВнешнееОснованиеДляРабот.Ссылка В(&МассивДокОсн)
	|
	|ОБЪЕДИНИТЬ ВСЕ
	|
	|ВЫБРАТЬ
	|	торо_ВыявленныеДефекты.Ссылка,
	|	торо_ВыявленныеДефекты.Проведен
	|ИЗ
	|	Документ.торо_ВыявленныеДефекты КАК торо_ВыявленныеДефекты
	|ГДЕ
	|	торо_ВыявленныеДефекты.Ссылка В(&МассивДокОсн)
	|
	|ОБЪЕДИНИТЬ ВСЕ
	|
	|ВЫБРАТЬ
	|	торо_ЗаявкаНаРемонт.Ссылка,
	|	торо_ЗаявкаНаРемонт.Проведен
	|ИЗ
	|	Документ.торо_ЗаявкаНаРемонт КАК торо_ЗаявкаНаРемонт
	|ГДЕ
	|	торо_ЗаявкаНаРемонт.Ссылка В(&МассивДокОсн)
	|
	|ОБЪЕДИНИТЬ ВСЕ
	|
	|ВЫБРАТЬ
	|	торо_ПланГрафикРемонта.Ссылка,
	|	торо_ПланГрафикРемонта.Проведен
	|ИЗ
	|	Документ.торо_ПланГрафикРемонта КАК торо_ПланГрафикРемонта
	|ГДЕ
	|	торо_ПланГрафикРемонта.Ссылка В(&МассивДокОсн)");

	Запрос.УстановитьПараметр("МассивДокОсн", МассивДокументовОснований); 
	Выборка = Запрос.Выполнить().Выбрать();

	Пока Выборка.Следующий() Цикл
		Если Не Выборка.Проведен Тогда
			ТекстСообщения = НСтр("ru = 'Есть непроведенный документ основание: " + Строка(Выборка.Ссылка) + "'");
			ОбщегоНазначения.СообщитьПользователю(ТекстСообщения,,,, Отказ);
			Возврат;
		КонецЕсли;		
	КонецЦикла;	

	УстановитьУправляемыеБлокировки();

	ШаблонСообщения = НСтр("ru = 'Для объекта ремонта ""%1"" с видом ремонта ""%2"" отсутствуют строки в дереве ремонтных работ.'");
	Для Каждого СтрокаСРемонтом Из РемонтыОборудования Цикл
		МассивСтрок = РемонтныеРаботы.НайтиСтроки(Новый структура("РемонтыОборудования_ID", СтрокаСРемонтом.ID));
		Если МассивСтрок.Количество() = 0 Тогда			
			ТекстСообщения = СтрШаблон(ШаблонСообщения, СтрокаСРемонтом.ОбъектРемонта, СтрокаСРемонтом.ВидРемонтныхРабот);
			ОбщегоНазначения.СообщитьПользователю(ТекстСообщения,,,,Отказ);
			Возврат;			
		КонецЕсли;
	КонецЦикла;

	Если ПолучитьФункциональнуюОпцию("торо_ИспользоватьНарядДопускНаРаботыПовышеннойОпасностиОбязательно") Тогда
		РемонтныеРаботыМассив = РемонтныеРаботы.ВыгрузитьКолонку("РемонтнаяРабота");

		Запрос = Новый Запрос;
		Запрос.Текст = "ВЫБРАТЬ
		|	торо_ТехнологическиеОперации.Ссылка КАК Ссылка
		|ИЗ
		|	Справочник.торо_ТехнологическиеОперации КАК торо_ТехнологическиеОперации
		|ГДЕ
		|	торо_ТехнологическиеОперации.РаботаПовышеннойОпасности = ИСТИНА
		|	И торо_ТехнологическиеОперации.Ссылка В(&РемонтныеРаботы)";

		Запрос.УстановитьПараметр("РемонтныеРаботы", РемонтныеРаботыМассив); 
		Выборка = Запрос.Выполнить().Выбрать();

		Если Выборка.Следующий() Тогда
			ТекстСообщения = НСтр("ru = 'Необходимо распечатать наряд-допуск, так как присутствуют работы повышенной опасности.'");
			ОбщегоНазначения.СообщитьПользователю(ТекстСообщения);
		КонецЕсли;
	КонецЕсли;

	Если ДокументыОснования.Количество()>0 Тогда
		ПроверитьДатуДокумента(Отказ);
	КонецЕсли;

	РемонтыОборудованияТаблицаЗначений = РемонтыОборудования.Выгрузить(); 
	РемонтыОтсутствующиеВДокументахИсточниках = торо_Ремонты.ПроверитьНаличиеРемонтовВДокументахИсточникахПоIDРемонта(РемонтыОборудованияТаблицаЗначений);

	Если НЕ РемонтыОтсутствующиеВДокументахИсточниках = Неопределено Тогда

		Для каждого Ремонт Из РемонтыОтсутствующиеВДокументахИсточниках Цикл
			ШаблонСообщения = НСтр("ru = 'Для объекта ремонта ""%1"" отсутствует соответствующий ремонт в документе основании ""%2"".'");
			ТекстСообщения = СтрШаблон(ШаблонСообщения, Ремонт.ОбъектРемонта, Ремонт.ДокументИсточник);
			ОбщегоНазначения.СообщитьПользователю(ТекстСообщения,,,,Отказ);			
		КонецЦикла;

	КонецЕсли;

	Если Не Отказ Тогда
		ДвиженияПоРегистрам(РежимПроведения, Отказ);
	КонецЕсли;

	торо_РаботаСоСтатусамиДокументовСервер.УстановитьСтатусРемонтовПриПроведении(ЭтотОбъект);
	торо_РаботаСоСтатусамиДокументовСервер.УстановитьСтатусДокумента(Ссылка, Ссылка, Перечисления.торо_СтатусыДокументов.Зарегистрирован);
	торо_РаботаСоСтатусамиДокументовСервер.ИзменитьСтатусыДокументовРемонта(Ссылка);

	торо_Ремонты.ОбновитьЗаписиНезависимыхРегистровПоРемонтам(ЭтотОбъект, РежимЗаписиДокумента.Проведение);
	
	#Вставка
	//++ Проф-ИТ, #222, Башинская А.Ю., 03.08.2023, создание док.Состояние ОР на основании 
	проф_СоздатьСостояниеОбъектовРемонтаНаОсновании();
	//-- Проф-ИТ, #222, Башинская А.Ю., 03.08.2023  
	
	//++ Проф-ИТ, #66, Башинская А.Ю., 14.09.2023
	проф_ЗарегистрироватьИзмененияСметыДляИзмененияСтатусаВЕРП();
	//-- Проф-ИТ, #66, Башинская А.Ю., 14.09.2023  
	#КонецВставки 

КонецПроцедуры     

&ИзменениеИКонтроль("ОбработкаЗаполнения")
Процедура проф_ОбработкаЗаполнения(Основание, СтандартнаяОбработка) 			
	Если ТипЗнч(Основание) = Тип("Структура") Тогда
		торо_ЗаполнениеДокументов20.ЗаполнитьШапкуДокумента(Основание, ЭтотОбъект);	
	
		// Вызывается из формы операций без нарядов, актуально для опасных работ.
		
		Если Основание.Свойство("ОпасныеРаботы") И Основание.ОпасныеРаботы = Истина Тогда
			ВыводФормыПодбораПриОткрытии = Ложь;
		КонецЕсли;
		
		Если Основание.Свойство("Основание") И ОбщегоНазначения.ЭтоСсылка(ТипЗнч(Основание.Основание)) Тогда
			НоваяСтрока = ДокументыОснования.Добавить();
			НоваяСтрока.ДокументОснование = Основание.Основание;
		КонецЕсли;
		
		Если Не Основание.Свойство("ВыбранныеОперации") Тогда
			Возврат;
		КонецЕсли;
		
		ЗаполнитьТЧПриОпасныхРаботах(Основание);
	ИначеЕсли ТипЗнч(Основание) = Тип("Соответствие") Тогда
		
		торо_ЗаполнениеДокументов20.ЗаполнитьШапкуДокумента(Основание, ЭтотОбъект);	
		торо_ЗаполнениеДокументов20.ЗаполнитьДокументНаОсновании(Основание, ЭтотОбъект);
		
	ИначеЕсли Основание <> Неопределено Тогда
		
		Отказ = Ложь;
		торо_ЗаполнениеДокументов.ПроверитьВозможностьВводаНаОсновании(Основание, Отказ, Ложь);
		Если Отказ Тогда
			Возврат;
		КонецЕсли;
		
		торо_ЗаполнениеДокументов20.ЗаполнитьШапкуДокумента(Основание, ЭтотОбъект);	
		
		ИмяТЧ = торо_Ремонты.ПолучитьИмяТЧРемонтов(Основание);
		
		Неотмененный = торо_ЗаполнениеДокументов.ВернутьIDНеотмененногоРемонта(Основание, Основание[ИмяТЧ]);
		Если ИмяТЧ <> Неопределено И Неотмененный <> "" Тогда
			СоответствиеИДДокументам = Новый Соответствие;
			СоответствиеИДДокументам.Вставить(Неотмененный, Основание);   
			торо_ЗаполнениеДокументов20.ЗаполнитьДокументНаОсновании(СоответствиеИДДокументам, ЭтотОбъект);
		Иначе
			торо_ЗаполнениеДокументов20.ЗаполнитьДокументыОснования(Основание, ЭтотОбъект);
		КонецЕсли;
		#Вставка 
		//++ Проф-ИТ, #250, Карпов.Д.Ю., 20.09.2023
		Если ТипЗнч(Основание) = Тип("ДокументСсылка.торо_ЗаявкаНаРемонт") Тогда
			РемонтыОборудования.Загрузить(Основание.РемонтыОборудования.Выгрузить());
			РемонтныеРаботы.Загрузить(Основание.РемонтныеРаботы.Выгрузить());
			Исполнители.Загрузить(Основание.ТрудовыеЗатраты.Выгрузить());
			ИсполнителиПоРемонтам.Загрузить(Основание.ИсполнителиПоРемонтам.Выгрузить());
		КонецЕсли; 
		//-- Проф-ИТ, #250, Карпов.Д.Ю., 20.09.2023
		#КонецВставки
	Иначе
		торо_ЗаполнениеДокументов20.ЗаполнитьШапкуДокумента(Основание, ЭтотОбъект);	
	КонецЕсли;
КонецПроцедуры

#КонецОбласти

#Область СлужебныеПроцедурыИФункции

//++ Проф-ИТ, #222, Башинская А.Ю., 03.08.2023, создание док.Состояние ОР на основании
Процедура проф_СоздатьСостояниеОбъектовРемонтаНаОсновании()
	
	НайденныйДокумент = проф_НайтиСозданныйРанееДокументСостояниеОР();
	
	Если НайденныйДокумент = Неопределено Тогда
		// Нужно создавать документ
		ДокументСОР = Документы.торо_СостоянияОбъектовРемонта.СоздатьДокумент();
		ДокументСОР.Дата = ТекущаяДатаСеанса();
	Иначе  
		// Обновим уже созданный документ
		ДокументСОР = НайденныйДокумент.ПолучитьОбъект(); 
		ДокументСОР.ОбъектыРемонта.Очистить();
	КонецЕсли;
	
	ДокументСОР.Заполнить(Ссылка);
	ДокументСОР.Автор = Ссылка.Автор;
	ДокументСОР.Ответственный = Ссылка.Ответственный;	
	ДокументСОР.Записать(РежимЗаписиДокумента.Проведение); 
	
КонецПроцедуры

Функция проф_НайтиСозданныйРанееДокументСостояниеОР()
	
	Запрос = Новый Запрос;
	Запрос.Текст = 
	"ВЫБРАТЬ
	|	торо_СостоянияОбъектовРемонта.Ссылка КАК Ссылка
	|ИЗ
	|	Документ.торо_СостоянияОбъектовРемонта КАК торо_СостоянияОбъектовРемонта
	|ГДЕ
	|	НЕ торо_СостоянияОбъектовРемонта.ПометкаУдаления       
	|	И торо_СостоянияОбъектовРемонта.Проведен
	|	И торо_СостоянияОбъектовРемонта.ДокументОснование = &ДокументОснование";
	Запрос.УстановитьПараметр("ДокументОснование", Ссылка);
	РезультатЗапроса = Запрос.Выполнить();	
	ВыборкаДетальныеЗаписи = РезультатЗапроса.Выбрать();
	
	Пока ВыборкаДетальныеЗаписи.Следующий() Цикл
		Возврат	ВыборкаДетальныеЗаписи.Ссылка;
	КонецЦикла;
	
	Возврат Неопределено;
	
КонецФункции
//-- Проф-ИТ, #222, Башинская А.Ю., 03.08.2023 

//++ Проф-ИТ, #66, Башинская А.Ю., 14.09.2023
Процедура проф_ЗарегистрироватьИзмененияСметыДляИзмененияСтатусаВЕРП()
	
	Запрос = Новый Запрос;
	Запрос.Текст = 
	"ВЫБРАТЬ
	|	торо_НарядНаВыполнениеРемонтныхРаботРемонтыОборудования.ID КАК ID,
	|	торо_ЗаявкиПоРемонтам.Регистратор КАК Смета
	|ИЗ
	|	Документ.торо_НарядНаВыполнениеРемонтныхРабот.РемонтыОборудования КАК торо_НарядНаВыполнениеРемонтныхРаботРемонтыОборудования
	|		ЛЕВОЕ СОЕДИНЕНИЕ РегистрСведений.торо_ЗаявкиПоРемонтам КАК торо_ЗаявкиПоРемонтам
	|		ПО торо_НарядНаВыполнениеРемонтныхРаботРемонтыОборудования.ID = торо_ЗаявкиПоРемонтам.IDРемонта
	|ГДЕ
	|	торо_НарядНаВыполнениеРемонтныхРаботРемонтыОборудования.Ссылка = &Ссылка";	
	Запрос.УстановитьПараметр("Ссылка", Ссылка);	
	РезультатЗапроса = Запрос.Выполнить();
	
	ВыборкаДетальныеЗаписи = РезультатЗапроса.Выбрать();
	Заявка = Неопределено;	
	Пока ВыборкаДетальныеЗаписи.Следующий() Цикл  
		Если ЗначениеЗаполнено(ВыборкаДетальныеЗаписи.Смета) Тогда
			Заявка = ВыборкаДетальныеЗаписи.Смета;
		КонецЕсли;
	КонецЦикла;
	
	Если Заявка = Неопределено Тогда 
		Возврат;
	КонецЕсли;
		
    ЗаявкаОбъект = Заявка.получитьОбъект();  
	ЗаявкаОбъект.ОбменДанными.Загрузка = Истина;
	ИмяПланаОбмена = "ОбменТОИР30ЕРП20";
	Если Заявка.Проведен Тогда
		РежимЗаписи = РежимЗаписиДокумента.Проведение;
	Иначе
		РежимЗаписи = РежимЗаписиДокумента.Запись;
	КонецЕсли;
		
	торо_ОбменТОИР30ЕРП20События.ОбменТОиР30ЕРП20_ЗарегистрироватьИзменениеДокументаПередЗаписью(ЗаявкаОбъект, 
		Ложь, РежимЗаписи, РежимПроведенияДокумента.Неоперативный);
	Попытка
		ЗаявкаОбъект.Записать();
	Исключение
		ОписаниеОшибки = ОбработкаОшибок.ПодробноеПредставлениеОшибки(ИнформацияОбОшибке());
		ЗаписьЖурналаРегистрации(
			"Документ.торо_НарядНаВыполнениеРемонтныхРабот.проф_ЗарегистрироватьИзмененияСметыДляИзмененияСтатусаВЕРП",
			УровеньЖурналаРегистрации.Ошибка,
			Метаданные.Документы.торо_НарядНаВыполнениеРемонтныхРабот,
			Ссылка, 
			ОписаниеОшибки);
	КонецПопытки;
	
КонецПроцедуры

&После("ПередЗаписью")
Процедура проф_ПередЗаписью(Отказ, РежимЗаписи, РежимПроведения)
	
	Если ОбменДанными.Загрузка Тогда
		Возврат;
	КонецЕсли;
	
	//++ Проф-ИТ, #401, Соловьев А.А., 11.12.2023
	Если РежимЗаписи = РежимЗаписиДокумента.Проведение Тогда 
		проф_ОбщегоНазначенияВызовСервера.ПроверитьПризнакПодразделенияОрганизации(ЭтотОбъект["Подразделение"], Отказ);
	КонецЕсли;
	//-- Проф-ИТ, #401, Соловьев А.А., 11.12.2023
	
КонецПроцедуры
//-- Проф-ИТ, #66, Башинская А.Ю., 14.09.2023

#КонецОбласти

