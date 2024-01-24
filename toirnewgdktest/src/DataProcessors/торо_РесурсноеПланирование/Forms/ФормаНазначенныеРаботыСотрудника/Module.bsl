#Область ОбработчикиСобытийФормы

&НаСервере
Процедура ПриСозданииНаСервере(Отказ, СтандартнаяОбработка)
	
	Если Параметры.Свойство("ДатаГрафика") Тогда
		Дата = Параметры.ДатаГрафика;
	КонецЕсли;
	
	Если Параметры.Свойство("Сотрудник") Тогда
		Сотрудник = Параметры.Сотрудник;
	КонецЕсли;
	
	ОтображатьПлан = Истина;
	ЗаполнитьИнформациюПоРаботам();
		
КонецПроцедуры
	
#КонецОбласти

#Область СлужебныеПроцедурыИФункции

#Область Данные

&НаСервере
Функция ПолучитьИнформациюОПлановыхОперациях()
	
	СтруктураВозврата = Новый Структура("ВремяРаботы, Количество", 0, 0);
	
	Запрос = Новый Запрос;
	Запрос.Текст = 
	"ВЫБРАТЬ РАЗРЕШЕННЫЕ
	|	торо_НарядНаВыполнениеРемонтныхРаботИсполнители.Ссылка КАК Ссылка,
	|	торо_НарядНаВыполнениеРемонтныхРаботИсполнители.ID КАК ID,
	|	торо_НарядНаВыполнениеРемонтныхРаботИсполнители.Квалификация КАК Квалификация,
	|	торо_НарядНаВыполнениеРемонтныхРаботИсполнители.ДатаНачала КАК ДатаНачала,
	|	торо_НарядНаВыполнениеРемонтныхРаботИсполнители.Сотрудник КАК Сотрудник,
	|	торо_НарядНаВыполнениеРемонтныхРаботИсполнители.РемонтыОборудования_ID КАК Работа_ID,
	|	торо_НарядНаВыполнениеРемонтныхРаботИсполнители.ВремяРаботы КАК ВремяРаботы
	|ПОМЕСТИТЬ втДанные
	|ИЗ
	|	Документ.торо_НарядНаВыполнениеРемонтныхРабот.Исполнители КАК торо_НарядНаВыполнениеРемонтныхРаботИсполнители
	|ГДЕ
	|	НАЧАЛОПЕРИОДА(торо_НарядНаВыполнениеРемонтныхРаботИсполнители.ДатаНачала, ДЕНЬ) = &ДатаНачала
	|	И торо_НарядНаВыполнениеРемонтныхРаботИсполнители.Сотрудник = &Сотрудник
	|   И торо_НарядНаВыполнениеРемонтныхРаботИсполнители.Ссылка.Проведен
	|	И НЕ торо_НарядНаВыполнениеРемонтныхРаботИсполнители.Ссылка.ПометкаУдаления
	|;
	|
	|////////////////////////////////////////////////////////////////////////////////
	|ВЫБРАТЬ
	|	КОЛИЧЕСТВО(втДанные.ID) КАК Количество,
	|	втДанные.Сотрудник КАК Сотрудник,
	|	СУММА(втДанные.ВремяРаботы) КАК ВремяРаботы
	|ИЗ
	|	втДанные КАК втДанные
	|
	|СГРУППИРОВАТЬ ПО
	|	втДанные.Сотрудник";
	
	Запрос.УстановитьПараметр("ДатаНачала", Дата);
	Запрос.УстановитьПараметр("Сотрудник", Сотрудник);
	
	Если УчитыватьНепроведенныеДокументы Тогда
		Запрос.Текст = СтрЗаменить(Запрос.Текст, "И торо_НарядНаВыполнениеРемонтныхРаботИсполнители.Ссылка.Проведен", "");
	КонецЕсли;
	
	МассивСтатусов = Новый Массив;
	МассивСтатусов.Добавить(Перечисления.торо_СтатусыДокументов.ВРаботе);
	МассивСтатусов.Добавить(Перечисления.торо_СтатусыДокументов.Выполнен);
	
	Запрос.УстановитьПараметр("Статусы", МассивСтатусов);
	
	РезультатЗапроса = Запрос.Выполнить();
	
	ВыборкаДетальныеЗаписи = РезультатЗапроса.Выбрать();
	
	Пока ВыборкаДетальныеЗаписи.Следующий() Цикл
		ЗаполнитьЗначенияСвойств(СтруктураВозврата, ВыборкаДетальныеЗаписи);
	КонецЦикла;
	
	Возврат СтруктураВозврата;
	
КонецФункции

&НаСервере
Функция ПолучитьИнформациюОФактическихОперациях()

	СтруктураВозврата = Новый Структура("ВремяРаботы, Количество", 0, 0);
	
	Запрос = Новый Запрос;
	Запрос.Текст = 
		"ВЫБРАТЬ РАЗРЕШЕННЫЕ
		|	торо_АктОВыполненииЭтапаРаботТрудовыеЗатраты.Ссылка КАК Ссылка,
		|	торо_АктОВыполненииЭтапаРаботТрудовыеЗатраты.ID КАК ID,
		|	торо_АктОВыполненииЭтапаРаботТрудовыеЗатраты.Сотрудник КАК Сотрудник,
		|	торо_АктОВыполненииЭтапаРаботРемонтныеРаботы.Выполнено КАК Выполнено,
		|	торо_АктОВыполненииЭтапаРаботТрудовыеЗатраты.РемонтыОборудования_ID КАК Работа_ID,
		|	торо_АктОВыполненииЭтапаРаботТрудовыеЗатраты.ДатаНачала КАК ДатаНачалаФакт,
		|	торо_АктОВыполненииЭтапаРаботТрудовыеЗатраты.ВремяРаботы КАК ВремяРаботыФакт,
		|	NULL КАК Статус
		|ПОМЕСТИТЬ втДанные
		|ИЗ
		|	Документ.торо_АктОВыполненииЭтапаРабот.ТрудовыеЗатраты КАК торо_АктОВыполненииЭтапаРаботТрудовыеЗатраты
		|		ЛЕВОЕ СОЕДИНЕНИЕ Документ.торо_АктОВыполненииЭтапаРабот.РемонтныеРаботы КАК торо_АктОВыполненииЭтапаРаботРемонтныеРаботы
		|		ПО торо_АктОВыполненииЭтапаРаботТрудовыеЗатраты.Ссылка = торо_АктОВыполненииЭтапаРаботРемонтныеРаботы.Ссылка
		|			И торо_АктОВыполненииЭтапаРаботТрудовыеЗатраты.ID = торо_АктОВыполненииЭтапаРаботРемонтныеРаботы.ID
		|			И торо_АктОВыполненииЭтапаРаботТрудовыеЗатраты.РемонтыОборудования_ID = торо_АктОВыполненииЭтапаРаботРемонтныеРаботы.РемонтыОборудования_ID
		|ГДЕ
		|	торо_АктОВыполненииЭтапаРаботТрудовыеЗатраты.Ссылка.Проведен
		|	И НЕ торо_АктОВыполненииЭтапаРаботТрудовыеЗатраты.Ссылка.ПометкаУдаления
		|	И торо_АктОВыполненииЭтапаРаботТрудовыеЗатраты.Сотрудник = &Сотрудник
		|	И НАЧАЛОПЕРИОДА(торо_АктОВыполненииЭтапаРаботТрудовыеЗатраты.ДатаНачала, ДЕНЬ) = &ДатаНачала
		|	И торо_АктОВыполненииЭтапаРаботРемонтныеРаботы.Выполнено
		|;
		|
		|////////////////////////////////////////////////////////////////////////////////
		|ВЫБРАТЬ
		|	КОЛИЧЕСТВО(втДанные.ID) КАК Количество,
		|	втДанные.Сотрудник КАК Сотрудник,
		|	СУММА(втДанные.ВремяРаботыФакт) КАК ВремяРаботы
		|ИЗ
		|	втДанные КАК втДанные
		|
		|СГРУППИРОВАТЬ ПО
		|	втДанные.Сотрудник";
	
	Запрос.УстановитьПараметр("ДатаНачала", Дата);
	Запрос.УстановитьПараметр("Сотрудник", Сотрудник);
	Запрос.УстановитьПараметр("Статус", Перечисления.торо_СтатусыДокументов.Выполнен);
		
	РезультатЗапроса = Запрос.Выполнить();
	
	ВыборкаДетальныеЗаписи = РезультатЗапроса.Выбрать();
	
	Пока ВыборкаДетальныеЗаписи.Следующий() Цикл
		ЗаполнитьЗначенияСвойств(СтруктураВозврата, ВыборкаДетальныеЗаписи);
	КонецЦикла;
	
	Возврат СтруктураВозврата;
	
КонецФункции

&НаСервере
Функция ПолучитьПлановуюДоступностьСотрудника()
	
	ПлановаяДоступность = 0;
	
	Запрос = Новый Запрос;
	Запрос.Текст = 
		"ВЫБРАТЬ
		|	ЕСТЬNULL(СУММА(торо_СостояниеРаботыСотрудниковСрезПоследних.ВремяРаботы), 0) КАК ВремяРаботы
		|ИЗ
		|	РегистрСведений.торо_СостояниеРаботыСотрудников.СрезПоследних КАК торо_СостояниеРаботыСотрудниковСрезПоследних
		|ГДЕ
		|	НЕ торо_СостояниеРаботыСотрудниковСрезПоследних.Отменен
		|	И торо_СостояниеРаботыСотрудниковСрезПоследних.ДатаГрафика = &ДатаГрафика
		|	И торо_СостояниеРаботыСотрудниковСрезПоследних.Сотрудник = &Сотрудник
		|	И торо_СостояниеРаботыСотрудниковСрезПоследних.СостояниеСотрудника = &СостояниеСотрудника";
	
	Запрос.УстановитьПараметр("ДатаГрафика", Дата);
	Запрос.УстановитьПараметр("Сотрудник", Сотрудник);
	Запрос.УстановитьПараметр("СостояниеСотрудника", Справочники.торо_СостоянияСотрудника.Явка);
	
	РезультатЗапроса = Запрос.Выполнить();
	
	ВыборкаДетальныеЗаписи = РезультатЗапроса.Выбрать();
	
	Пока ВыборкаДетальныеЗаписи.Следующий() Цикл
		ПлановаяДоступность = ВыборкаДетальныеЗаписи.ВремяРаботы;	
	КонецЦикла;
	
	Если ПлановаяДоступность > 0 Тогда
		ПлановаяДоступность = ПолучитьВремяВЧасах(ПлановаяДоступность);
	КонецЕсли;	
	
	Возврат ПлановаяДоступность;
	
КонецФункции

&НаСервере
Функция ПолучитьТаблицуРабот()
	
	Если ОтображатьПлан Тогда
		
		Запрос = Новый Запрос;
		Запрос.Текст = 
		"ВЫБРАТЬ РАЗРЕШЕННЫЕ
		|	торо_НарядНаВыполнениеРемонтныхРаботИсполнители.Ссылка КАК Ссылка,
		|	торо_НарядНаВыполнениеРемонтныхРаботИсполнители.ДатаНачала КАК ДатаНачала,
		|	торо_НарядНаВыполнениеРемонтныхРаботИсполнители.ДатаОкончания КАК ДатаОкончания,
		|	торо_НарядНаВыполнениеРемонтныхРаботИсполнители.ID КАК ID,
		|	торо_НарядНаВыполнениеРемонтныхРаботИсполнители.РемонтыОборудования_ID КАК Работа_ID,
		|	торо_НарядНаВыполнениеРемонтныхРаботИсполнители.Сотрудник КАК Сотрудник,
		|	РАЗНОСТЬДАТ(торо_НарядНаВыполнениеРемонтныхРаботИсполнители.ДатаНачала, торо_НарядНаВыполнениеРемонтныхРаботИсполнители.ДатаОкончания, СЕКУНДА) КАК ВремяРаботы,
		|	торо_НарядНаВыполнениеРемонтныхРаботРемонтыОборудования.ОбъектРемонта КАК ОбъектРемонта,
		|	торо_НарядНаВыполнениеРемонтныхРаботРемонтыОборудования.ВидРемонтныхРабот КАК ВидРемонтныхРабот,
		|	торо_НарядНаВыполнениеРемонтныхРаботРемонтныеРаботы.РемонтнаяРабота КАК Операция,
		|	ЛОЖЬ КАК Выполнено
		|ИЗ
		|	Документ.торо_НарядНаВыполнениеРемонтныхРабот.Исполнители КАК торо_НарядНаВыполнениеРемонтныхРаботИсполнители
		|		ЛЕВОЕ СОЕДИНЕНИЕ Документ.торо_НарядНаВыполнениеРемонтныхРабот.РемонтныеРаботы КАК торо_НарядНаВыполнениеРемонтныхРаботРемонтныеРаботы
		|			ЛЕВОЕ СОЕДИНЕНИЕ Документ.торо_НарядНаВыполнениеРемонтныхРабот.РемонтыОборудования КАК торо_НарядНаВыполнениеРемонтныхРаботРемонтыОборудования
		|			ПО торо_НарядНаВыполнениеРемонтныхРаботРемонтныеРаботы.Ссылка = торо_НарядНаВыполнениеРемонтныхРаботРемонтыОборудования.Ссылка
		|				И торо_НарядНаВыполнениеРемонтныхРаботРемонтныеРаботы.РемонтыОборудования_ID = торо_НарядНаВыполнениеРемонтныхРаботРемонтыОборудования.ID
		|		ПО торо_НарядНаВыполнениеРемонтныхРаботИсполнители.Ссылка = торо_НарядНаВыполнениеРемонтныхРаботРемонтныеРаботы.Ссылка
		|			И торо_НарядНаВыполнениеРемонтныхРаботИсполнители.ID = торо_НарядНаВыполнениеРемонтныхРаботРемонтныеРаботы.ID
		|			И торо_НарядНаВыполнениеРемонтныхРаботИсполнители.РемонтыОборудования_ID = торо_НарядНаВыполнениеРемонтныхРаботРемонтныеРаботы.РемонтыОборудования_ID
		|ГДЕ
		|	НАЧАЛОПЕРИОДА(торо_НарядНаВыполнениеРемонтныхРаботИсполнители.ДатаНачала, ДЕНЬ) = &ДатаНачала
		|	И торо_НарядНаВыполнениеРемонтныхРаботИсполнители.Сотрудник = &Сотрудник
		|	И НЕ торо_НарядНаВыполнениеРемонтныхРаботИсполнители.Ссылка.ПометкаУдаления
		|	И (&УчитыватьНепроведенные
		|			ИЛИ торо_НарядНаВыполнениеРемонтныхРаботИсполнители.Ссылка.Проведен)";
				
		МассивСтатусов = Новый Массив;
		МассивСтатусов.Добавить(Перечисления.торо_СтатусыДокументов.ВРаботе);
		МассивСтатусов.Добавить(Перечисления.торо_СтатусыДокументов.Выполнен);
		
		Запрос.УстановитьПараметр("Статусы", МассивСтатусов);
		
	Иначе
		
		Запрос = Новый Запрос;
		Запрос.Текст = 
		"ВЫБРАТЬ РАЗРЕШЕННЫЕ
		|	торо_АктОВыполненииЭтапаРаботТрудовыеЗатраты.Ссылка КАК Ссылка,
		|	торо_АктОВыполненииЭтапаРаботТрудовыеЗатраты.ID КАК ID,
		|	торо_АктОВыполненииЭтапаРаботТрудовыеЗатраты.Сотрудник КАК Сотрудник,
		|	торо_АктОВыполненииЭтапаРаботРемонтныеРаботы.Выполнено КАК Выполнено,
		|	торо_АктОВыполненииЭтапаРаботТрудовыеЗатраты.РемонтыОборудования_ID КАК Работа_ID,
		|	торо_АктОВыполненииЭтапаРаботТрудовыеЗатраты.ДатаНачала КАК ДатаНачала,
		|	торо_АктОВыполненииЭтапаРаботТрудовыеЗатраты.ВремяРаботы КАК ВремяРаботы,
		|	торо_АктОВыполненииЭтапаРаботТрудовыеЗатраты.ДатаОкончания КАК ДатаОкончания,
		|	торо_АктОВыполненииЭтапаРаботРемонтныеРаботы.РемонтнаяРабота КАК Операция,
		|	торо_АктОВыполненииЭтапаРаботРемонтыОборудования.ОбъектРемонта КАК ОбъектРемонта,
		|	торо_АктОВыполненииЭтапаРаботРемонтыОборудования.ВидРемонтныхРабот КАК ВидРемонтныхРабот
		|ИЗ
		|	Документ.торо_АктОВыполненииЭтапаРабот.ТрудовыеЗатраты КАК торо_АктОВыполненииЭтапаРаботТрудовыеЗатраты
		|		ЛЕВОЕ СОЕДИНЕНИЕ Документ.торо_АктОВыполненииЭтапаРабот.РемонтныеРаботы КАК торо_АктОВыполненииЭтапаРаботРемонтныеРаботы
		|			ЛЕВОЕ СОЕДИНЕНИЕ Документ.торо_АктОВыполненииЭтапаРабот.РемонтыОборудования КАК торо_АктОВыполненииЭтапаРаботРемонтыОборудования
		|			ПО торо_АктОВыполненииЭтапаРаботРемонтныеРаботы.Ссылка = торо_АктОВыполненииЭтапаРаботРемонтыОборудования.Ссылка
		|				И торо_АктОВыполненииЭтапаРаботРемонтныеРаботы.РемонтыОборудования_ID = торо_АктОВыполненииЭтапаРаботРемонтыОборудования.ID
		|		ПО торо_АктОВыполненииЭтапаРаботТрудовыеЗатраты.Ссылка = торо_АктОВыполненииЭтапаРаботРемонтныеРаботы.Ссылка
		|			И торо_АктОВыполненииЭтапаРаботТрудовыеЗатраты.ID = торо_АктОВыполненииЭтапаРаботРемонтныеРаботы.ID
		|			И торо_АктОВыполненииЭтапаРаботТрудовыеЗатраты.РемонтыОборудования_ID = торо_АктОВыполненииЭтапаРаботРемонтныеРаботы.РемонтыОборудования_ID
		|ГДЕ
		|	(&УчитыватьНепроведенные
		|			ИЛИ торо_АктОВыполненииЭтапаРаботТрудовыеЗатраты.Ссылка.Проведен)
		|	И НЕ торо_АктОВыполненииЭтапаРаботТрудовыеЗатраты.Ссылка.ПометкаУдаления
		|	И торо_АктОВыполненииЭтапаРаботТрудовыеЗатраты.Сотрудник = &Сотрудник
		|	И НАЧАЛОПЕРИОДА(торо_АктОВыполненииЭтапаРаботТрудовыеЗатраты.ДатаНачала, ДЕНЬ) = &ДатаНачала
		|	И торо_АктОВыполненииЭтапаРаботРемонтныеРаботы.Выполнено";
		
		Запрос.УстановитьПараметр("Статус", Перечисления.торо_СтатусыДокументов.Выполнен);
		
	КонецЕсли;
	
	Запрос.УстановитьПараметр("ДатаНачала", Дата);
	Запрос.УстановитьПараметр("Сотрудник", Сотрудник);
	Запрос.УстановитьПараметр("УчитыватьНепроведенные", УчитыватьНепроведенныеДокументы);
	
	Выгрузка = Запрос.Выполнить().Выгрузить();
	
	Возврат Выгрузка;
	
КонецФункции

#КонецОбласти

&НаСервере
Процедура ЗаполнитьИнформациюПоРаботам()

	ИнформацияПоРаботам.Очистить();
		
	НС = ИнформацияПоРаботам.Добавить();
	
	ОперацииПлан = ПолучитьИнформациюОПлановыхОперациях();
	ОперацииФакт = ПолучитьИнформациюОФактическихОперациях();

	НС.КоличествоОперацийПлан 	= ОперацииПлан.Количество;
	НС.КоличествоОперацийФакт 	= ОперацииФакт.Количество;
	НС.Табель 					= ПолучитьПлановуюДоступностьСотрудника();
	НС.Назначено 				= ПолучитьВремяВЧасах(ОперацииПлан.ВремяРаботы);
	НС.Свободно 				= НС.Табель - НС.Назначено;
	НС.ЗакрытоАктами 			= ПолучитьВремяВЧасах(ОперацииФакт.ВремяРаботы);
	
	ЗаполнитьДиаграмму();
	
КонецПроцедуры

&НаСервере
Процедура ЗаполнитьДиаграмму()

	ДиаграммаГанта.Очистить();
	ДиаграммаГанта.Обновление = Ложь;
	
	ДиаграммаГанта.УстановитьПолныйИнтервал(НачалоДня(Дата), КонецДня(Дата));
	
	ТаблицаРабот = ПолучитьТаблицуРабот();
	
	Если ТаблицаРабот.Количество() > 0 Тогда
		
		лТаблицаРабот = ТаблицаРабот.Скопировать(,"ОбъектРемонта, ВидРемонтныхРабот, Работа_ID");
		лТаблицаРабот.Свернуть("ОбъектРемонта, ВидРемонтныхРабот, Работа_ID");
		
		Для каждого ТекСтрока Из лТаблицаРабот Цикл
			
			СтруктураСтроки = Новый Структура("ОбъектРемонта, ВидРемонтныхРабот, Работа_ID");
			ЗаполнитьЗначенияСвойств(СтруктураСтроки, ТекСтрока);
			
			ТекСерия 					= ДиаграммаГанта.УстановитьСерию(ТекСтрока.Работа_ID);
			ТекСерия.ДополнительныйЦвет = WebЦвета.СветлоСерый;
			
			ТочкаРемонта 					= ДиаграммаГанта.УстановитьТочку(СтруктураСтроки);
			ТочкаРемонта.Картинка 			= БиблиотекаКартинок.торо_ДобавитьРемонтнуюРаботу;
			ТочкаРемонта.ДополнительныйЦвет = WebЦвета.СветлоСерый;
			ТочкаРемонта.Текст = СтрШаблон("%1 %2", торо_ЗаполнениеДокументов.ПолучитьПредоставленияВРДляПечати(СтруктураСтроки.ВидРемонтныхРабот), СтруктураСтроки.ОбъектРемонта);
						
			ДанныеРемонта = ТаблицаРабот.НайтиСтроки(СтруктураСтроки);
			
			Для каждого СтрокаДанных Из ДанныеРемонта Цикл
				
				ЗначениеРемонта = ДиаграммаГанта.ПолучитьЗначение(ТочкаРемонта, ТекСерия);
				
				ИнтервалРемонта 		= ЗначениеРемонта.Добавить();
				ИнтервалРемонта.Начало 	= СтрокаДанных.ДатаНачала;
				ИнтервалРемонта.Конец 	= СтрокаДанных.ДатаОкончания;
				ИнтервалРемонта.Цвет 	= WebЦвета.СветлоСерый;
				
				СтруктураТочки = Новый Структура("Документ", СтрокаДанных.Ссылка);
				
				ТочкаРаботы 			= ДиаграммаГанта.УстановитьТочку(СтруктураТочки, СтруктураСтроки);
				ТочкаРаботы.Текст 		= СтрокаДанных.Операция;
				ТочкаРаботы.Картинка 	= БиблиотекаКартинок.торо_ТехОперация;
				
				Значение = ДиаграммаГанта.ПолучитьЗначение(ТочкаРаботы, ТекСерия);
				
				Интервал 		= Значение.Добавить();
				Интервал.Начало = СтрокаДанных.ДатаНачала;
				Интервал.Конец 	= СтрокаДанных.ДатаОкончания;
				Интервал.Цвет 	= ?(СтрокаДанных.Выполнено, WebЦвета.ЛимонноЗеленый, WebЦвета.СинеСерый);				
				Интервал.Текст 	= СтрШаблон("%1 - %2" + Символы.ПС + "<%3>", Формат(СтрокаДанных.ДатаНачала, "ДЛФ=T"), Формат(СтрокаДанных.ДатаОкончания, "ДЛФ=T"), торо_Ремонты.ПродолжительностьВЧасах(СтрокаДанных.ВремяРаботы)); 
				
			КонецЦикла;
			
		КонецЦикла;
		
	КонецЕсли;
	
	ОтобразитьИнтервалыРаботыСотрудника();
	
	Для каждого Точка Из ДиаграммаГанта.Точки Цикл
		ДиаграммаГанта.РазвернутьТочку(Точка,Ложь);
	КонецЦикла; 
	
	ДиаграммаГанта.Обновление = Истина;
	
КонецПроцедуры

&НаСервере
Процедура ОтобразитьИнтервалыРаботыСотрудника()
	
	Запрос = Новый Запрос("ВЫБРАТЬ
	                      |	торо_СостояниеРаботыСотрудниковСрезПоследних.ВремяНачала КАК НачалоИнтервала,
	                      |	торо_СостояниеРаботыСотрудниковСрезПоследних.ВремяОкончания КАК КонецИнтервала
	                      |ИЗ
	                      |	РегистрСведений.торо_СостояниеРаботыСотрудников.СрезПоследних КАК торо_СостояниеРаботыСотрудниковСрезПоследних
	                      |ГДЕ
	                      |	торо_СостояниеРаботыСотрудниковСрезПоследних.Сотрудник = &Сотрудник
	                      |	И НЕ торо_СостояниеРаботыСотрудниковСрезПоследних.Отменен
	                      |	И торо_СостояниеРаботыСотрудниковСрезПоследних.СостояниеСотрудника = ЗНАЧЕНИЕ(Справочник.торо_СостоянияСотрудника.Явка)
	                      |	И торо_СостояниеРаботыСотрудниковСрезПоследних.ДатаГрафика = &ДатаНачала");
	Запрос.УстановитьПараметр("Сотрудник", Сотрудник);
	Запрос.УстановитьПараметр("ДатаНачала", Дата);
	
	Выборка = Запрос.Выполнить().Выбрать();
	Пока Выборка.Следующий() Цикл
		
		ИФ = ДиаграммаГанта.ИнтервалыФона.Добавить(Мин(Выборка.НачалоИнтервала, Выборка.КонецИнтервала), Макс(Выборка.НачалоИнтервала, Выборка.КонецИнтервала));
		ИФ.Цвет = WebЦвета.АкварельноСиний;
	
	КонецЦикла; 

КонецПроцедуры

&НаСервере
Функция ПолучитьВремяВЧасах(ВремяВСекундах = 0)
	
	ВремяВЧасах = 0;
	
	Если ВремяВСекундах > 0 Тогда
		ВремяВЧасах = ВремяВСекундах / 60 /60;	
	КонецЕсли;
	
	Возврат ВремяВЧасах;
	
КонецФункции

#КонецОбласти

#Область ОбработчикиСобытийЭлементовШапкиФормы

&НаКлиенте
Процедура ОтображатьФактПриИзменении(Элемент)
	ЗаполнитьДиаграмму();
КонецПроцедуры

&НаКлиенте
Процедура УчитыватьНепроведенныеДокументыПриИзменении(Элемент)
	ЗаполнитьИнформациюПоРаботам();	
КонецПроцедуры

&НаКлиенте
Процедура ДиаграммаГантаВыбор(Элемент, Значения, СтандартнаяОбработка, Дата)
	
	СтандартнаяОбработка = Ложь;
	
	Если ТипЗнч(Значения) = Тип("ТочкаДиаграммыГанта") Тогда
		
		ЗначениеТочки = Значения.Значение;
		Если ТипЗнч(ЗначениеТочки) = Тип("Структура") И ЗначениеТочки.Свойство("Документ") Тогда
			ПоказатьЗначение(,ЗначениеТочки.Документ);
		КонецЕсли;
		
	ИначеЕсли ТипЗнч(Значения) = Тип("Массив") Тогда
		
		Для каждого Элт Из Значения Цикл
			Если ТипЗнч(Элт) = Тип("ЗначениеДиаграммыГанта") Тогда
				ЗначениеТочки = Элт.Точка.Значение;
				Если ТипЗнч(ЗначениеТочки) = Тип("Структура") И ЗначениеТочки.Свойство("Документ") Тогда
					ПоказатьЗначение(,ЗначениеТочки.Документ);
				КонецЕсли; 
			КонецЕсли;
		КонецЦикла; 
		
	КонецЕсли;
	
КонецПроцедуры

#КонецОбласти

