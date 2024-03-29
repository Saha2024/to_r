
#Область ОбработчикиСобытийФормы

&НаСервере
Процедура ПриСозданииНаСервере(Отказ, СтандартнаяОбработка)
	
	Если Параметры.Свойство("АвтоТест") Тогда
		
		Возврат;
		
	КонецЕсли;
	
	ЗапретитьСоздание = Константы.торо_ЗапретитьВводЗаказовНаВнутреннееПотреблениеБезЗаявки.Получить();
	
	Если ЗапретитьСоздание Тогда 
		Если Элементы.Найти("ФормаСоздать") <> Неопределено Тогда
			Элементы.ФормаСоздать.Доступность = Ложь; 
		КонецЕсли;	
		Если Элементы.Найти("СписокКонтекстноеМенюСоздать") <> Неопределено Тогда
			Элементы.СписокКонтекстноеМенюСоздать.Доступность = Ложь;
			Элементы.СписокКонтекстноеМенюСоздать.Видимость = Ложь;
		КонецЕсли;
		Если Элементы.Найти("ФормаСкопировать") <> Неопределено Тогда
			Элементы.ФормаСкопировать.Доступность = Ложь;   
		КонецЕсли;
		Если Элементы.Найти("СписокКонтекстноеМенюСкопировать") <> Неопределено Тогда
			Элементы.СписокКонтекстноеМенюСкопировать.Доступность = Ложь;
			Элементы.СписокКонтекстноеМенюСкопировать.Видимость = Ложь;
		КонецЕсли;
	КонецЕсли;
	
	// СтандартныеПодсистемы.ПодключаемыеКоманды
	ПараметрыРазмещения = ПодключаемыеКоманды.ПараметрыРазмещения();
	ПараметрыРазмещения.КоманднаяПанель = Элементы.ГруппаГлобальныеКоманды;
	ПодключаемыеКоманды.ПриСозданииНаСервере(ЭтаФорма, ПараметрыРазмещения);
	// Конец СтандартныеПодсистемы.ПодключаемыеКоманды
	
	// СтандартныеПодсистемы.ВерсионированиеОбъектов
	ВерсионированиеОбъектов.ПриСозданииНаСервере(ЭтотОбъект);
	// Конец СтандартныеПодсистемы.ВерсионированиеОбъектов
	
	ИспользоватьВнутреннееПотреблениеПоНесколькимЗаказам = Истина;
	торо_СобытияФорм.ПриСозданииНаСервере(ЭтаФорма, Отказ, СтандартнаяОбработка, Новый Структура("УстановитьСвойствоЭлементовФормыОтПрав", Истина));
	
	Если Параметры.Свойство("торо_СозданПоПотребности") Тогда
		ЭлементОтбора = Список.Отбор.Элементы.Добавить(Тип("ЭлементОтбораКомпоновкиДанных"));
		ЭлементОтбора.ЛевоеЗначение = Новый ПолеКомпоновкиДанных("торо_СозданПоПотребности");
		ЭлементОтбора.ВидСравнения = ВидСравненияКомпоновкиДанных.Равно;
		ЭлементОтбора.Использование = Истина;
		ЭлементОтбора.ПравоеЗначение = Параметры.торо_СозданПоПотребности;
	КонецЕсли; 
	
	Если Параметры.Свойство("ДатаОтгрузки") Тогда
		ЭлементОтбора = Список.Отбор.Элементы.Добавить(Тип("ЭлементОтбораКомпоновкиДанных"));
		ЭлементОтбора.ЛевоеЗначение = Новый ПолеКомпоновкиДанных("ДатаОтгрузки");
		ЭлементОтбора.ВидСравнения = ВидСравненияКомпоновкиДанных.БольшеИлиРавно;
		ЭлементОтбора.Использование = Истина;
		ЭлементОтбора.ПравоеЗначение = Параметры.ДатаОтгрузки;
	КонецЕсли;
	
	Элементы.торо_СтатусУтверждения.Видимость = Константы.торо_ИспользоватьУтверждениеЗаказовНаВнутреннееПотребление.Получить();

КонецПроцедуры

#КонецОбласти

#Область ОбработчикиКомандФормы

&НаКлиенте
Процедура СоздатьВнутреннееПотреблениеТоваров(Команда)
	
	СоздатьВнутреннееПотреблениеТоваровНаОснованииЗаказа();
	
КонецПроцедуры

// СтандартныеПодсистемы.ПодключаемыеКоманды
&НаКлиенте
Процедура Подключаемый_ВыполнитьКоманду(Команда)
	ПодключаемыеКомандыКлиент.НачатьВыполнениеКоманды(ЭтотОбъект, Команда, Элементы.Список);
КонецПроцедуры
 
&НаКлиенте
Процедура Подключаемый_ПродолжитьВыполнениеКомандыНаСервере(ПараметрыВыполнения, ДополнительныеПараметры) Экспорт
	ВыполнитьКомандуНаСервере(ПараметрыВыполнения);
КонецПроцедуры
 
&НаСервере
Процедура ВыполнитьКомандуНаСервере(ПараметрыВыполнения)
	ПодключаемыеКоманды.ВыполнитьКоманду(ЭтотОбъект, ПараметрыВыполнения, Элементы.Список);
КонецПроцедуры
 
&НаКлиенте
Процедура Подключаемый_ОбновитьКоманды()
	ПодключаемыеКомандыКлиентСервер.ОбновитьКоманды(ЭтотОбъект, Элементы.Список);
КонецПроцедуры
// Конец СтандартныеПодсистемы.ПодключаемыеКоманды

#КонецОбласти

#Область ОбработчикиСобытийЭлементовТаблицыФормыСписок

&НаКлиенте
Процедура СписокПриАктивизацииСтроки(Элемент)
	
	// СтандартныеПодсистемы.ПодключаемыеКоманды
	ПодключаемыеКомандыКлиент.НачатьОбновлениеКоманд(ЭтотОбъект);
	// Конец СтандартныеПодсистемы.ПодключаемыеКоманды
	
	Если Элементы.Список.ТекущиеДанные <> Неопределено Тогда 
		ДоступностьИзменения = НЕ Элементы.Список.ТекущиеДанные.торо_СозданПоПотребности; 
		Если Элементы.Найти("СписокКонтекстноеМенюУстановитьПометкуУдаления") <> Неопределено Тогда
			Элементы.СписокКонтекстноеМенюУстановитьПометкуУдаления.Доступность = ДоступностьИзменения; 
		КонецЕсли; 
		Если Элементы.Найти("ФормаУстановитьПометкуУдаления") <> Неопределено Тогда
			Элементы.ФормаУстановитьПометкуУдаления.Доступность = ДоступностьИзменения; 
		КонецЕсли; 
		Если Элементы.Найти("Проведение") <> Неопределено Тогда
			Элементы.Проведение.Доступность = ДоступностьИзменения; 
		КонецЕсли; 
		Если Элементы.Найти("ФормаПроведение") <> Неопределено Тогда
			Элементы.ФормаПроведение.Доступность = ДоступностьИзменения; 
		КонецЕсли;
	КонецЕсли;

КонецПроцедуры

#КонецОбласти

#Область СлужебныеПроцедурыИФункции

#Область Прочее

&НаКлиенте
Процедура СоздатьВнутреннееПотреблениеТоваровНаОснованииЗаказа()
	
	ТекущиеДанные = Элементы.Список.ТекущиеДанные;
	
	Если ТекущиеДанные = Неопределено Тогда
		
		ТекстПредупреждения = НСтр("ru = 'Команда не может быть выполнена для указанного объекта!'");
		ПоказатьПредупреждение(Неопределено, ТекстПредупреждения);
		Возврат;
		
	КонецЕсли;
	
	РеквизитыШапки = Новый Структура();
	МассивСсылок = Новый Массив();
	
	Если Элементы.Список.ВыделенныеСтроки.Количество() = 1 Или Не ИспользоватьВнутреннееПотреблениеПоНесколькимЗаказам Тогда
		
		Если ТипЗнч(Элементы.Список.ТекущаяСтрока) = Тип("СтрокаГруппировкиДинамическогоСписка") Тогда
			
			ТекстПредупреждения = НСтр("ru = 'Команда не может быть выполнена для указанного объекта!'");
			ПоказатьПредупреждение(Неопределено, ТекстПредупреждения);
			Возврат;
			
		КонецЕсли;
		
		ПараметрыОснования = Новый Структура;
		ПараметрыОснования.Вставить("ДокументОснование", ТекущиеДанные.Ссылка);
		ОткрытьФорму("Документ.ВнутреннееПотреблениеТоваров.ФормаОбъекта",
			Новый Структура("Основание", ПараметрыОснования));
		
	Иначе
		
		Для каждого Заказ из Элементы.Список.ВыделенныеСтроки Цикл
			
			Если ТипЗнч(Заказ) = Тип("СтрокаГруппировкиДинамическогоСписка") Тогда
				
				Продолжить;
				
			КонецЕсли;
			
			МассивСсылок.Добавить(Заказ);
			
		КонецЦикла;
		
		Если МассивСсылок.Количество() = 0 Тогда
			
			ТекстПредупреждения = НСтр("ru = 'Команда не может быть выполнена для указанного объекта!'");
			ПоказатьПредупреждение(Неопределено, ТекстПредупреждения);
			Возврат;
			
		КонецЕсли;
		
		ТекстПредупреждения = Неопределено;
		РеквизитыШапки = ДанныеЗаполненияВнутреннегоПотребления(МассивСсылок, ТекстПредупреждения);
		Если РеквизитыШапки = Неопределено Тогда
			Если ТекстПредупреждения <> Неопределено Тогда
				ПоказатьПредупреждение(, ТекстПредупреждения);
			КонецЕсли; 
			Возврат
		КонецЕсли;
			
		ПараметрыОснования = Новый Структура;
		ПараметрыОснования.Вставить("РеквизитыШапки", РеквизитыШапки);
		ПараметрыОснования.Вставить("ДокументОснование", МассивСсылок);
		
		ОткрытьФорму("Документ.ВнутреннееПотреблениеТоваров.ФормаОбъекта",
			Новый Структура("Основание", ПараметрыОснования));
			
		
	КонецЕсли;
	
КонецПроцедуры

&НаСервере
Функция ДанныеЗаполненияВнутреннегоПотребления(Знач МассивСсылок, ТекстПредупреждения)
	
	Возврат Документы.ЗаказНаВнутреннееПотребление.ПараметрыОформленияВнутреннегоПотребления(МассивСсылок, ТекстПредупреждения);
	
КонецФункции

#КонецОбласти

#КонецОбласти
