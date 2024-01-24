///////////////////////////////////////////////////////////////////////////////////////////////////////
// Copyright (c) 2022, ООО 1С-Софт
// Все права защищены. Эта программа и сопроводительные материалы предоставляются 
// в соответствии с условиями лицензии Attribution 4.0 International (CC BY 4.0)
// Текст лицензии доступен по ссылке:
// https://creativecommons.org/licenses/by/4.0/legalcode
///////////////////////////////////////////////////////////////////////////////////////////////////////

#Если Сервер Или ТолстыйКлиентОбычноеПриложение Или ВнешнееСоединение Тогда

#Область ПрограммныйИнтерфейс

#Область ДляВызоваИзДругихПодсистем

// СтандартныеПодсистемы.ГрупповоеИзменениеОбъектов

// Возвращает реквизиты объекта, которые разрешается редактировать
// с помощью обработки группового изменения реквизитов.
//
// Возвращаемое значение:
//  Массив из Строка
//
Функция РеквизитыРедактируемыеВГрупповойОбработке() Экспорт
	
	Результат = Новый Массив;
	Возврат Результат;
	
КонецФункции

// Конец СтандартныеПодсистемы.ГрупповоеИзменениеОбъектов

#КонецОбласти

#КонецОбласти

#Область СлужебныеПроцедурыИФункции

// Возвращает ссылку на справочник внешней компоненты по идентификатору и версии.
//
// Параметры:
//  Идентификатор - Строка - идентификатор объекта внешнего компонента.
//  Версия        - Строка - версия компоненты.
//
// Возвращаемое значение:
//  СправочникСсылка.ВнешниеКомпоненты - ссылка на контейнер внешней компоненты в информационной базе.
//
Функция НайтиПоИдентификатору(Идентификатор, Версия = Неопределено) Экспорт
	
	Запрос = Новый Запрос;
	Запрос.УстановитьПараметр("Идентификатор", Идентификатор);
	
	Если Не ЗначениеЗаполнено(Версия) Тогда 
		Запрос.Текст = 
			"ВЫБРАТЬ ПЕРВЫЕ 1
			|	ВнешниеКомпоненты.Идентификатор КАК Идентификатор,
			|	ВнешниеКомпоненты.ДатаВерсии КАК ДатаВерсии,
			|	ВЫБОР
			|		КОГДА ВнешниеКомпоненты.Использование = ЗНАЧЕНИЕ(Перечисление.ВариантыИспользованияВнешнихКомпонент.Используется)
			|			ТОГДА ИСТИНА
			|		ИНАЧЕ ЛОЖЬ
			|	КОНЕЦ КАК Использование,
			|	ВнешниеКомпоненты.Ссылка КАК Ссылка
			|ИЗ
			|	Справочник.ВнешниеКомпоненты КАК ВнешниеКомпоненты
			|ГДЕ
			|	ВнешниеКомпоненты.Идентификатор = &Идентификатор
			|
			|УПОРЯДОЧИТЬ ПО
			|	Использование УБЫВ,
			|	ДатаВерсии УБЫВ";
	Иначе 
		Запрос.УстановитьПараметр("Версия", Версия);
		Запрос.Текст = 
			"ВЫБРАТЬ ПЕРВЫЕ 1
			|	ВнешниеКомпоненты.Ссылка КАК Ссылка,
			|	ВЫБОР
			|		КОГДА ВнешниеКомпоненты.Использование = ЗНАЧЕНИЕ(Перечисление.ВариантыИспользованияВнешнихКомпонент.Используется)
			|			ТОГДА ИСТИНА
			|		ИНАЧЕ ЛОЖЬ
			|	КОНЕЦ КАК Использование
			|ИЗ
			|	Справочник.ВнешниеКомпоненты КАК ВнешниеКомпоненты
			|ГДЕ
			|	ВнешниеКомпоненты.Идентификатор = &Идентификатор
			|	И ВнешниеКомпоненты.Версия = &Версия
			|
			|УПОРЯДОЧИТЬ ПО
			|	Использование УБЫВ";
		
	КонецЕсли;
	
	Результат = Запрос.Выполнить();
	
	Если Результат.Пустой() Тогда 
		Возврат ПустаяСсылка();
	КонецЕсли;
	
	Выборка = Результат.Выбрать();
	Выборка.Следующий();
	
	Возврат Результат.Выгрузить()[0].Ссылка;
	
КонецФункции

#Область ОбработчикиОбновления

// Регистрирует на плане обмена ОбновлениеИнформационнойБазы объекты,
// которые необходимо обновить на новую версию.
//
Процедура ЗарегистрироватьДанныеКОбработкеДляПереходаНаНовуюВерсию(Параметры) Экспорт
	
	ТекстЗапроса ="ВЫБРАТЬ
	|	ВнешниеКомпоненты.Ссылка
	|ИЗ
	|	Справочник.ВнешниеКомпоненты КАК ВнешниеКомпоненты";
	
	Запрос = Новый Запрос(ТекстЗапроса);
	
	ОбновлениеИнформационнойБазы.ОтметитьКОбработке(Параметры, Запрос.Выполнить().Выгрузить().ВыгрузитьКолонку("Ссылка"));
	
КонецПроцедуры

// Обработчик обновления на версию 3.1.6.48:
// - заполняет реквизиты совместимости для браузеров MacOS в справочнике "Внешние компоненты".
//
Процедура ОбработатьДанныеДляПереходаНаНовуюВерсию(Параметры) Экспорт
	
	Выборка = ОбновлениеИнформационнойБазы.ВыбратьСсылкиДляОбработки(Параметры.Очередь,
		"Справочник.ВнешниеКомпоненты");
	Если Выборка.Количество() > 0 Тогда
		ОбработатьВнешниеКомпоненты(Выборка);
	КонецЕсли;

	ОбработкаЗавершена = ОбновлениеИнформационнойБазы.ОбработкаДанныхЗавершена(Параметры.Очередь,
		"Справочник.ВнешниеКомпоненты");
	Параметры.ОбработкаЗавершена = ОбработкаЗавершена;
			
КонецПроцедуры

Процедура ОбработатьВнешниеКомпоненты(Выборка)
	
	ОбъектовОбработано = 0;
	ПроблемныхОбъектов = 0;

	Пока Выборка.Следующий() Цикл

		РеквизитыКомпоненты = ОбщегоНазначения.ЗначенияРеквизитовОбъекта(Выборка.Ссылка,
			"ХранилищеКомпоненты, MacOS_x86_64_Safari, MacOS_x86_64_Chrome, MacOS_x86_64_Firefox");
			
		Если ТипЗнч(РеквизитыКомпоненты.ХранилищеКомпоненты) <> Тип("ХранилищеЗначения") Тогда
			ОбновлениеИнформационнойБазы.ОтметитьВыполнениеОбработки(Выборка.Ссылка);
			ОбъектовОбработано = ОбъектовОбработано + 1;
			Продолжить;
		КонецЕсли;
			
		ДвоичныеДанныеКомпоненты = РеквизитыКомпоненты.ХранилищеКомпоненты.Получить();
		
		Если ТипЗнч(ДвоичныеДанныеКомпоненты) <> Тип("ДвоичныеДанные") Тогда
			ОбновлениеИнформационнойБазы.ОтметитьВыполнениеОбработки(Выборка.Ссылка);
			ОбъектовОбработано = ОбъектовОбработано + 1;
			Продолжить;
		КонецЕсли;
		
		ИнформацияОКомпонентеИзФайла = ВнешниеКомпонентыСлужебный.ИнформацияОКомпонентеИзФайла(
			ДвоичныеДанныеКомпоненты, Ложь);
		Если Не ИнформацияОКомпонентеИзФайла.Разобрано Тогда
			ОбновлениеИнформационнойБазы.ОтметитьВыполнениеОбработки(Выборка.Ссылка);
			ОбъектовОбработано = ОбъектовОбработано + 1;
			Продолжить;
		КонецЕсли;
		
		Реквизиты = ИнформацияОКомпонентеИзФайла.Реквизиты;
		
		Если РеквизитыКомпоненты.MacOS_x86_64_Safari = ?(Реквизиты.MacOS_x86_64_Safari = Неопределено, Ложь,
				Реквизиты.MacOS_x86_64_Safari)
			И РеквизитыКомпоненты.MacOS_x86_64_Chrome = ?(Реквизиты.MacOS_x86_64_Chrome = Неопределено, Ложь,
				Реквизиты.MacOS_x86_64_Chrome)
			И РеквизитыКомпоненты.MacOS_x86_64_Firefox = ?(Реквизиты.MacOS_x86_64_Firefox = Неопределено, Ложь,
				Реквизиты.MacOS_x86_64_Firefox) Тогда
			
			ОбновлениеИнформационнойБазы.ОтметитьВыполнениеОбработки(Выборка.Ссылка);
			ОбъектовОбработано = ОбъектовОбработано + 1;
			
			Продолжить;
		КонецЕсли;
		
		НачатьТранзакцию();
		Попытка

			Блокировка = Новый БлокировкаДанных;
			ЭлементБлокировки = Блокировка.Добавить("Справочник.ВнешниеКомпоненты");
			ЭлементБлокировки.УстановитьЗначение("Ссылка", Выборка.Ссылка);
			ЭлементБлокировки.Режим = РежимБлокировкиДанных.Разделяемый;
			Блокировка.Заблокировать();

			КомпонентаОбъект = Выборка.Ссылка.ПолучитьОбъект(); // СправочникОбъект.ВнешниеКомпоненты
			КомпонентаОбъект.MacOS_x86_64_Safari = Реквизиты.MacOS_x86_64_Safari;
			КомпонентаОбъект.MacOS_x86_64_Chrome = Реквизиты.MacOS_x86_64_Chrome;
			КомпонентаОбъект.MacOS_x86_64_Firefox = Реквизиты.MacOS_x86_64_Firefox;
			ОбновлениеИнформационнойБазы.ЗаписатьОбъект(КомпонентаОбъект);

			ОбъектовОбработано = ОбъектовОбработано + 1;
			ЗафиксироватьТранзакцию();

		Исключение

			ОтменитьТранзакцию();
			// Если не удалось обработать компоненту, повторяем попытку снова.
			ПроблемныхОбъектов = ПроблемныхОбъектов + 1;

			ТекстСообщения = СтроковыеФункцииКлиентСервер.ПодставитьПараметрыВСтроку(
				НСтр("ru = 'Не удалось обработать компоненту %1 по причине:
					 |%2'"), Выборка.Ссылка, ПодробноеПредставлениеОшибки(ИнформацияОбОшибке()));

			ЗаписьЖурналаРегистрации(ОбновлениеИнформационнойБазы.СобытиеЖурналаРегистрации(),
				УровеньЖурналаРегистрации.Предупреждение, Выборка.Ссылка.Метаданные(), Выборка.Ссылка, ТекстСообщения);

		КонецПопытки;

	КонецЦикла;

	Если ОбъектовОбработано = 0 И ПроблемныхОбъектов <> 0 Тогда
		ТекстСообщения = СтроковыеФункцииКлиентСервер.ПодставитьПараметрыВСтроку(
			НСтр("ru = 'Не удалось обработать некоторые компоненты (пропущены): %1'"),
			ПроблемныхОбъектов);
		ВызватьИсключение ТекстСообщения;
	Иначе
		ЗаписьЖурналаРегистрации(ОбновлениеИнформационнойБазы.СобытиеЖурналаРегистрации(),
			УровеньЖурналаРегистрации.Информация, Метаданные.Справочники.ВнешниеКомпоненты,,
				СтроковыеФункцииКлиентСервер.ПодставитьПараметрыВСтроку(
					НСтр("ru = 'Обработана очередная порция компонент: %1'"),
			ОбъектовОбработано));
	КонецЕсли;
	
КонецПроцедуры

#КонецОбласти

#КонецОбласти

#КонецЕсли