#Область ПрограммныйИнтерфейс

// Выполнение подключаемой команды ЭДО.
//
// Параметры:
//  Команда - КомандаФормы - вызывающая команда.
//  Форма - ФормаКлиентскогоПриложения - вызывающая форма.
//  Источник - РеквизитФормы -реквизит формы.
//
Процедура ВыполнитьПодключаемуюКомандуЭДО(Знач Команда, Знач Форма, Знач Источник) Экспорт
	
	ОписаниеКоманды = Команда;
	Если ТипЗнч(Команда) = Тип("КомандаФормы") Тогда
		
		АдресКомандВоВременномХранилище = Форма.Команды.Найти("АдресКомандЭДОВоВременномХранилище").Действие;
		ОписаниеКоманды = ПодключаемыйКомандыЭДОСлужебныйВызовСервера.ОписаниеКомандыЭДО(Команда.Имя, 
																							АдресКомандВоВременномХранилище);
																							
	КонецЕсли;
	
	ДополнительныеПараметры = Новый Структура;
	ДополнительныеПараметры.Вставить("ОписаниеКоманды", ОписаниеКоманды);
	ДополнительныеПараметры.Вставить("Форма", Форма);
	ДополнительныеПараметры.Вставить("Источник", Источник);
	
	Если ТипЗнч(Источник) = Тип("ДанныеФормыСтруктура")
		И (Источник.Ссылка.Пустая() Или Форма.Модифицированность) Тогда
		
		ТекстВопроса = СтроковыеФункцииКлиентСервер.ПодставитьПараметрыВСтроку(
			НСтр("ru = 'Данные еще не записаны.
				|Выполнение действия ""%1"" возможно только после записи данных.
				|Данные будут записаны.'"),
			ОписаниеКоманды.Представление);
			
		ОписаниеОповещения = Новый ОписаниеОповещения("ВыполнитьПодключаемуюКомандуЭДОПодтверждениеЗаписи",
			ПодключаемыеКомандыЭДОСлужебныйКлиент, ДополнительныеПараметры);
		ПоказатьВопрос(ОписаниеОповещения, ТекстВопроса, РежимДиалогаВопрос.ОКОтмена);
		Возврат;
	КонецЕсли;
	
	ПодключаемыеКомандыЭДОСлужебныйКлиент.ВыполнитьПодключаемуюКомандуЭДОПодтверждениеЗаписи(
		Неопределено, ДополнительныеПараметры);
	
КонецПроцедуры

// Обновляет список команд в зависимости от текущего контекста.
//
// Параметры:
//   Форма - ФормаКлиентскогоПриложения
//   Источник - ДанныеФормыСтруктура, ТаблицаФормы - контекст для проверки условий (Форма.Объект или Форма.Элементы.Список)
//
Процедура ОбновитьКоманды(Форма, Источник) Экспорт

	Если ОбщегоНазначенияКлиентСервер.ЕстьРеквизитИлиСвойствоОбъекта(Форма, "ПараметрыУправленияВидимостьюЭДО") Тогда
		ПараметрыУправленияВидимостью = Форма.ПараметрыУправленияВидимостьюЭДО;
		Если ТипЗнч(ПараметрыУправленияВидимостью) <> Тип("Структура")
			Или Не ПараметрыУправленияВидимостью.ЕстьУсловияВидимости Тогда
			Возврат;
		КонецЕсли;
	Иначе
		Возврат;
	КонецЕсли;

	Если ТипЗнч(Источник) = Тип("ТаблицаФормы") Тогда
		Если ПараметрыУправленияВидимостью.УсловияВидимостиВТаблицеФормы Тогда
			ПодключаемыеКомандыЭДОСлужебныйКлиент.ОбновитьКомандыСпискаНаКлиенте(Форма, ПараметрыУправленияВидимостью, Источник);
		Иначе
			ПодключаемыеКомандыЭДОСлужебныйКлиент.ОбновитьКомандыСписка(Форма, ПараметрыУправленияВидимостью, Источник);
		КонецЕсли;
	Иначе
		ПодключаемыеКомандыЭДОСлужебныйКлиент.ОбновитьКомандыОбъекта(Форма, ПараметрыУправленияВидимостью, Источник);
	КонецЕсли;

КонецПроцедуры

#КонецОбласти
