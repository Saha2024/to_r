#Область ОбработчикиСобытийФормы

&НаСервере
Процедура ПриСозданииНаСервере(Отказ, СтандартнаяОбработка)
	
	МассивВыделенныхСтрок = ?(Параметры.Свойство("МассивВыделенныхСтрок"), Параметры.МассивВыделенныхСтрок, Новый Массив);
	СоответствиеРемонтовИОснований = ?(Параметры.Свойство("СоответствиеРемонтовИОснований"), Параметры.СоответствиеРемонтовИОснований, Новый Соответствие);
	
	Для каждого Строка из МассивВыделенныхСтрок Цикл
		
		МассивОснований = СоответствиеРемонтовИОснований[Строка.ID_Ремонта];
		Если МассивОснований = Неопределено ИЛИ МассивОснований.Количество() = 0 Тогда
			Продолжить;
		КонецЕсли;
		
		СтрокаДерева = ДеревоРемонтов.ПолучитьЭлементы().Добавить();
		СтрокаДерева.ОбъектРемонта = Строка.ОбъектРемонта;
		СтрокаДерева.ВидРемонта    = Строка.ВидРемонта;
		СтрокаДерева.ДатаРемонта   = Строка.ПланируемаяДатаНачалаРемонта;
		СтрокаДерева.ID            = Строка.ID_Ремонта;
		СтрокаДерева.Уровень       = 0;

		Индекс = МассивОснований.ВГраница();
        ПоследнийДокумент = МассивОснований [Индекс];
		Для каждого Основание из МассивОснований Цикл
			СтрокаВторогоУровня = СтрокаДерева.ПолучитьЭлементы().Добавить();
			СтрокаВторогоУровня.Документ     = Основание;
			СтрокаВторогоУровня.ID           = Строка.ID_Ремонта;
			СтрокаВторогоУровня.Уровень      = 1;

			Если Основание = ПоследнийДокумент Тогда
				СтрокаВторогоУровня.Выбран = Истина;
			КонецЕсли;
		КонецЦикла;
		
	КонецЦикла;
		
КонецПроцедуры

&НаКлиенте
Процедура ПриОткрытии(Отказ)
	
	ЭлементыДерева = ДеревоРемонтов.ПолучитьЭлементы();
	
	Для каждого Элемент Из ЭлементыДерева Цикл
		Элементы.ДеревоРемонтов.Развернуть(Элемент.ПолучитьИдентификатор(),Истина); 
	КонецЦикла; 
	
КонецПроцедуры

#КонецОбласти

#Область ОбработчикиСобытийЭлементовТаблицыФормыДеревоРемонтов

&НаКлиенте
Процедура ДеревоРемонтовВыбранПриИзменении(Элемент)
	
	ТекСтрокаДерева = Элементы.ДеревоРемонтов.ТекущиеДанные;
	
	ТекДокумент = ?(ТекСтрокаДерева.Выбран, ТекСтрокаДерева.Документ, Неопределено);
	
	Родитель = ТекСтрокаДерева.ПолучитьРодителя();
	
	Если Не ТекДокумент = Неопределено Тогда
		Для каждого ЭлементДерева Из Родитель.ПолучитьЭлементы() Цикл
			 Если Не ЭлементДерева.Документ = ТекДокумент Тогда
			 	 ЭлементДерева.Выбран = Ложь;
			 КонецЕсли; 
		КонецЦикла;
	Иначе
		Родитель.Выбран = Ложь;
	КонецЕсли;	
		
КонецПроцедуры

#КонецОбласти

#Область ОбработчикиКомандФормы

&НаКлиенте
Процедура ПередатьВыбранные(Команда)
	МассивРемонтов = Новый Массив;
	ЭлементыДерева = ДеревоРемонтов.ПолучитьЭлементы();
	
	Для каждого Элемент Из ЭлементыДерева Цикл
		
		ЭлементыДокументы = Элемент.ПолучитьЭлементы();
		
		Для каждого Строка Из ЭлементыДокументы Цикл
			Если Строка.Выбран И ЗначениеЗаполнено(Строка.Документ) Тогда
				Структура = Новый Структура("ID_Ремонта, Документ", Строка.ID,Строка.Документ);
				МассивРемонтов.Добавить(Структура);
			КонецЕсли; 
		КонецЦикла;
		
	КонецЦикла; 
		
	Закрыть(МассивРемонтов);
	
КонецПроцедуры

&НаКлиенте
Процедура СвернутьВсе(Команда)
	Для каждого Стр Из ДеревоРемонтов.ПолучитьЭлементы() Цикл
		
		СвернутьПодчиненные(Стр);
		Элементы.ДеревоРемонтов.Свернуть(Стр.ПолучитьИдентификатор());
		
	КонецЦикла;
КонецПроцедуры

&НаКлиенте
Процедура РазвернутьВсе(Команда)
	
	Для каждого Стр Из ДеревоРемонтов.ПолучитьЭлементы() Цикл
		Элементы.ДеревоРемонтов.Развернуть(Стр.ПолучитьИдентификатор(), Истина);
	КонецЦикла;
	
КонецПроцедуры

#КонецОбласти

#Область СлужебныеПроцедурыИФункции

&НаКлиенте
Процедура СвернутьПодчиненные(Строка)
	
	Для каждого Стр Из Строка.ПолучитьЭлементы() Цикл
		
		СвернутьПодчиненные(Стр);
		Элементы.ДеревоРемонтов.Свернуть(Стр.ПолучитьИдентификатор());
		
	КонецЦикла; 	
	
КонецПроцедуры

#КонецОбласти
 







