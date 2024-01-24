#Область ОбработчикиСобытийФормы
&НаСервере
Процедура ПриСозданииНаСервере(Отказ, СтандартнаяОбработка)
	Если Параметры.Свойство("ЭтоЗаключение") Тогда
		ЭтоЗаключение = Параметры.ЭтоЗаключение;
		Если ЭтоЗаключение Тогда
			ЭтаФорма.Заголовок = "Добавление заключения";
		Иначе
			ЭтаФорма.Заголовок = "Добавление причины";
		КонецЕсли; 
	Иначе
	    ЭтоЗаключение = Ложь;
		ЭтаФорма.Заголовок = "Добавление причины";
	КонецЕсли;
	
	Если Параметры.Свойство("Промежуточная") Тогда
		Промежуточная = Параметры.Промежуточная;
	Иначе
	    Промежуточная = Ложь;
	КонецЕсли;
	
	Если Параметры.Свойство("ДоИзменения") Тогда
		Если ТипЗнч(Параметры.ДоИзменения) = Тип("Строка") Тогда
			ПричинаСтрока = Параметры.ДоИзменения;	
		Иначе
			Коренная = Истина;
		КонецЕсли;
		Редактирование = Истина;
	КонецЕсли; 
КонецПроцедуры

&НаКлиенте
Процедура ПриОткрытии(Отказ)
	Элементы.Коренная.Видимость = НЕ ЭтоЗаключение И НЕ Промежуточная;
	Если Коренная = Истина Тогда
		Элементы.ПричинаСправочник.Видимость = Истина;
		Элементы.ПричинаСтрока.Видимость = Ложь;
	Иначе
		Элементы.ПричинаСправочник.Видимость = Ложь;
		Элементы.ПричинаСтрока.Видимость = Истина;
	КонецЕсли; 	
КонецПроцедуры
#КонецОбласти

#Область ОбработчикиСобытийЭлементовШапкиФормы
&НаКлиенте
Процедура КореннаяПриИзменении(Элемент)
	Если Коренная = Истина Тогда
		Элементы.ПричинаСправочник.Видимость = Истина;
		Элементы.ПричинаСтрока.Видимость = Ложь;
	Иначе
		Элементы.ПричинаСправочник.Видимость = Ложь;
		Элементы.ПричинаСтрока.Видимость = Истина;
	КонецЕсли; 
КонецПроцедуры

#КонецОбласти

#Область ОбработчикиСобытийЭлементовТаблицыФормыПричинаСправочник
&НаКлиенте
Процедура ПричинаСправочникВыбор(Элемент, ВыбраннаяСтрока, Поле, СтандартнаяОбработка)
	ТекущиеДанные = Элемент.ТекущиеДанные;
	Если ТекущиеДанные <> Неопределено И НЕ ТекущиеДанные.ЭтоГруппа Тогда
		Выбрать(Элемент);
	КонецЕсли; 
КонецПроцедуры

#КонецОбласти

#Область ОбработчикиКомандФормы
&НаКлиенте
Процедура Выбрать(Команда)
	
	Если Не Коренная И ПричинаСтрока = "" Тогда
		Если ЭтоЗаключение Тогда
			ОбщегоНазначенияКлиент.СообщитьПользователю(НСтр("ru = 'Введите заключение!'"));
		Иначе
			ОбщегоНазначенияКлиент.СообщитьПользователю(НСтр("ru = 'Введите причину!'"));
		КонецЕсли; 
		
	Иначе
		СтруктураВыбора = Новый Структура();
		СтруктураВыбора.Вставить("ЭтоЗаключение", ЭтоЗаключение);
		СтруктураВыбора.Вставить("Редактирование", Редактирование);
		Если Коренная И Элементы.ПричинаСправочник.ТекущиеДанные <> Неопределено Тогда
			Если НЕ Элементы.ПричинаСправочник.ТекущиеДанные.ЭтоГруппа Тогда
				СтруктураВыбора.Вставить("Причина", Элементы.ПричинаСправочник.ТекущиеДанные.Ссылка);
				ОповеститьОВыборе(СтруктураВыбора);
			Иначе
			    Оповещение = Новый ОписаниеОповещения("ПослеЗакрытияПредупреждение",
			        ЭтотОбъект);	
			    Текст = "Выберите элемент, а не группу!
				|Для раскрытия групп используйте ""Ctrl"" и стрелки вниз/вверх или клавиши ""+"" и ""-"" на дополнительной клавиатуре.";
			    ПоказатьПредупреждение(Оповещение, Текст, 0, "");	
			КонецЕсли; 
		Иначе
			СтруктураВыбора.Вставить("Причина", ПричинаСтрока);
			ОповеститьОВыборе(СтруктураВыбора);
		КонецЕсли; 
	КонецЕсли; 
	
КонецПроцедуры
 
&НаКлиенте
Процедура ПослеЗакрытияПредупреждение(Параметры) Экспорт	
	ТочкаОстанова = 0;
КонецПроцедуры

#КонецОбласти





