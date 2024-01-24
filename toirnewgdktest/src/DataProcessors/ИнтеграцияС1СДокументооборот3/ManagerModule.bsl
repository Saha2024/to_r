#Если Сервер Или ТолстыйКлиентОбычноеПриложение Или ВнешнееСоединение Тогда

#Область ПрограммныйИнтерфейс

// Выполняет поиск объектов ДО по реквизитам, сохраняя результат во временное хранилище
//
// Параметры:
//   ТипОбъекта - Строка - тип объекта, поиск по которому требуется выполнить
//   РеквизитыПоиска - Структура - описание реквизитов поиска
//   АдресВоВременномХранилище - Строка - параметр, в который помещается адрес результатов поиска
//   КоличествоРезультатов - Число - количество результатов поиска
//   ПредельноеКоличествоРезультатов - Число - предельное количество выбираемых объектов, если оно превышено
//
Процедура ВыполнитьПоискПоРеквизитам(ТипОбъекта, РеквизитыПоиска, АдресВоВременномХранилище,
		КоличествоРезультатов, ПредельноеКоличествоРезультатов) Экспорт
	
	Прокси = ИнтеграцияС1СДокументооборотБазоваяФункциональностьПовтИсп.ПолучитьПрокси();
	
	СписокУсловий = ИнтеграцияС1СДокументооборотБазоваяФункциональность.СоздатьОбъект(Прокси, "DMObjectListQuery");
	ПредельноеКоличествоРезультатов = ИнтеграцияС1СДокументооборотБазоваяФункциональность.ПредельноеКоличествоВыбираемыхОбъектов(
		ТипОбъекта);
	СписокУсловий.limit = ПредельноеКоличествоРезультатов;
	ВыбираемыеПоля = СписокУсловий.columnSet; // СписокXDTO
	
	НаборКолонок = Новый Массив;
	Если ТипОбъекта = "DMDocument" Тогда
		НаборКолонок.Добавить("documentType");
		НаборКолонок.Добавить("regNumber");
		НаборКолонок.Добавить("regDate");
		НаборКолонок.Добавить("sum");
		НаборКолонок.Добавить("organization");
		НаборКолонок.Добавить("correspondent");
		НаборКолонок.Добавить("folder");
		
	ИначеЕсли ТипОбъекта = "DMCorrespondent" Тогда
		НаборКолонок.Добавить("inn");
		НаборКолонок.Добавить("kpp");
		НаборКолонок.Добавить("legalPrivatePerson");
		
	ИначеЕсли ТипОбъекта = "DMMeeting" Тогда
		НаборКолонок.Добавить("type");
		НаборКолонок.Добавить("startDate");
		НаборКолонок.Добавить("endDate");
		
	КонецЕсли;
	
	Представление = "";
	Для Каждого РеквизитПоиска Из РеквизитыПоиска Цикл
		ИмяРеквизита = РеквизитПоиска.Ключ;
		ОписаниеУсловия = РеквизитПоиска.Значение;
		ДобавитьРеквизитВСписокУсловий(
			Прокси,
			ИмяРеквизита,
			ОписаниеУсловия,
			СписокУсловий,
			Представление);
		// Реквизиты с отбором на равенство не информативны.
		Если ОписаниеУсловия.ОператорСравнения = "=" Тогда
			ПозицияВМассиве = НаборКолонок.Найти(ИмяРеквизита);
			Если ПозицияВМассиве <> Неопределено Тогда
				НаборКолонок.Удалить(ПозицияВМассиве);
			КонецЕсли;
		КонецЕсли;
	КонецЦикла;
	
	// добавим оставшиеся колонки в требование выборки
	Для Каждого ИмяРеквизита Из НаборКолонок Цикл
		ВыбираемыеПоля.Добавить(ИмяРеквизита);
	КонецЦикла;
	
	Если Представление = "" Тогда
		Представление = НСтр("ru = 'Условия не заданы.'");
	Иначе
		Представление = Представление + ".";
	КонецЕсли;
	
	Результат = ИнтеграцияС1СДокументооборотБазоваяФункциональность.НайтиСписокОбъектов(
		Прокси,
		ТипОбъекта,
		СписокУсловий);
	
	НайденныеОбъекты = Новый Массив;
	Для Каждого ОбъектXDTO Из Результат.items Цикл
		НайденныйОбъект = Новый Структура;
		НайденныеОбъекты.Добавить(НайденныйОбъект);
		НайденныйОбъект.Вставить("name", ОбъектXDTO.object.name);
		НайденныйОбъект.Вставить("ID", ОбъектXDTO.object.objectID.ID);
		Для Каждого ОписаниеКолонки Из НаборКолонок Цикл
			ИмяРеквизита = ОписаниеКолонки;
			ЗначениеСвойства = ОбъектXDTO.object.Получить(ИмяРеквизита);
			Если ТипЗнч(ЗначениеСвойства) = Тип("ОбъектXDTO") Тогда // ссылочный тип ДО
				НайденныйОбъект.Вставить(ИмяРеквизита, ЗначениеСвойства.name);
				НайденныйОбъект.Вставить(ИмяРеквизита + "ID", ЗначениеСвойства.objectID.ID);
			Иначе // примитивный тип
				НайденныйОбъект.Вставить(ИмяРеквизита, ЗначениеСвойства);
			КонецЕсли;
		КонецЦикла;
	КонецЦикла;
	
	// ранние версии ДО не поддерживают tooManyObjects
	Если ?(Результат.Установлено("tooManyObjects"), Результат.tooManyObjects, Ложь) Тогда
		Представление = Представление
			+ " " + НСтр("ru = 'Перетащите сюда любой реквизит, чтобы уточнить условия.'");
		
	Иначе
		ПредельноеКоличествоРезультатов = 0;
	КонецЕсли;
	
	Результат = Новый Структура;
	Результат.Вставить("НайденныеОбъекты", НайденныеОбъекты);
	Результат.Вставить("НаборКолонок", НаборКолонок);
	Результат.Вставить("ПредставлениеУсловийПоиска", Представление);
	
	АдресВоВременномХранилище = ПоместитьВоВременноеХранилище(Результат, Новый УникальныйИдентификатор);
	
	КоличествоРезультатов = НайденныеОбъекты.Количество();
	
КонецПроцедуры

#КонецОбласти

#Область СлужебныеПроцедурыИФункции

// Добавляет реквизит поиска в список условий XDTO, параллельно формируя представление условия
//
// Параметры:
//   Прокси - WSПрокси - объект для подключения к web-сервисам Документооборота.
//   ИмяРеквизита - Строка - имя реквизита поиска.
//   ОписаниеУсловий - Структура - описание условий поиска.
//   СписокУсловий - ОбъектXDTO - ОбъектXDTO типа DMObjectListQuery, пополняемый список условий.
//   Представление - Строка - пополняемое представление условия.
//
Процедура ДобавитьРеквизитВСписокУсловий(Прокси, ИмяРеквизита, ОписаниеУсловия, СписокУсловий, Представление)
	
	Если Представление = "" Тогда
		Представление = Представление + ОписаниеУсловия.Представление;
	Иначе
		Представление = Представление + "; " + НРег(ОписаниеУсловия.Представление);
	КонецЕсли;
	
	Если ОписаниеУсловия.Свойство("ПредставлениеУсловия") И ЗначениеЗаполнено(ОписаниеУсловия.ПредставлениеУсловия) Тогда
		Представление = Представление + ": " + ОписаниеУсловия.ПредставлениеУсловия;
	Иначе
		Представление = Представление + ": "
			+ ?(ТипЗнч(ОписаниеУсловия.Значение) = Тип("Дата"),
				Формат(ОписаниеУсловия.Значение, "ДЛФ=D"),
				ОписаниеУсловия.Значение);
	КонецЕсли;
	
	Условие = ИнтеграцияС1СДокументооборотБазоваяФункциональность.СоздатьОбъект(Прокси, "DMObjectListCondition");
	Условие.property = ИмяРеквизита;
	Если ОписаниеУсловия.ОператорСравнения = "<=" И ТипЗнч(ОписаниеУсловия.Значение) = Тип("Дата") Тогда
		Условие.value = КонецДня(ОписаниеУсловия.Значение);
	Иначе
		Если ОписаниеУсловия.Свойство("ЗначениеID") И ЗначениеЗаполнено(ОписаниеУсловия.ЗначениеID) Тогда
			ЗначениеРеквизита = ИнтеграцияС1СДокументооборотБазоваяФункциональность.СоздатьObjectID(
				Прокси,
				ОписаниеУсловия.ЗначениеID,
				ОписаниеУсловия.ЗначениеТип);
			Условие.value = ЗначениеРеквизита;
		Иначе // примитивный тип
			Условие.value = ОписаниеУсловия.Значение;
		КонецЕсли;
	КонецЕсли;
	
	Если ОписаниеУсловия.ОператорСравнения = "=" И ТипЗнч(ОписаниеУсловия.Значение) = Тип("Дата") Тогда
		Условие.comparisonOperator = ">=";
	Иначе
		Условие.comparisonOperator = ОписаниеУсловия.ОператорСравнения;
	КонецЕсли;
	
	Условия = СписокУсловий.conditions; // СписокXDTO
	Условия.Добавить(Условие);
	
	Если ОписаниеУсловия.Свойство("ОператорСравнения2") И ЗначениеЗаполнено(ОписаниеУсловия.ОператорСравнения2) Тогда
		Условие = ИнтеграцияС1СДокументооборотБазоваяФункциональность.СоздатьОбъект(Прокси, "DMObjectListCondition");
		Условие.property = ИмяРеквизита;
		Если ОписаниеУсловия.ОператорСравнения2 = "<=" И ТипЗнч(ОписаниеУсловия.Значение2) = Тип("Дата") Тогда
			Условие.value = КонецДня(ОписаниеУсловия.Значение2);
		Иначе
			Условие.value = ОписаниеУсловия.Значение2;
		КонецЕсли;
		Условие.comparisonOperator = ОписаниеУсловия.ОператорСравнения2;
		Условия.Добавить(Условие);
	ИначеЕсли ОписаниеУсловия.ОператорСравнения = "="
			И ТипЗнч(ОписаниеУсловия.Значение) = Тип("Дата")
			И ИмяРеквизита <> "anyDate" Тогда
		Условие = ИнтеграцияС1СДокументооборотБазоваяФункциональность.СоздатьОбъект(Прокси, "DMObjectListCondition");
		Условие.property = ИмяРеквизита;
		Условие.value = КонецДня(ОписаниеУсловия.Значение);
		Условие.comparisonOperator = "<=";
		Условия.Добавить(Условие);
	КонецЕсли;
	
КонецПроцедуры

#КонецОбласти

#КонецЕсли