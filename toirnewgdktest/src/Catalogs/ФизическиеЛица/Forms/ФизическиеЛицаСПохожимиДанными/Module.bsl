
#Область ОбработчикиСобытийФормы

&НаСервере
Процедура ПриСозданииНаСервере(Отказ, СтандартнаяОбработка)
	
	Заголовок = Параметры.ЗаголовокФормы;
	ТекстИнформационнойНадписи = Параметры.ТекстИнформационнойНадписи;
	
	ДанныеФизическихЛиц.Очистить();
	
	Для Каждого ДанныеФизЛица Из Параметры.ДанныеФизическихЛиц Цикл
		
		НоваяСтрокаДанных = ДанныеФизическихЛиц.Добавить();
		ЗаполнитьЗначенияСвойств(НоваяСтрокаДанных, ДанныеФизЛица);
		
		Если НЕ ЗначениеЗаполнено(НоваяСтрокаДанных.ДатаРождения) Тогда
			НоваяСтрокаДанных.ДатаРождения = "<" + НСтр("ru='Не заполнена'") + ">";
		КонецЕсли;
		
		Если НЕ ЗначениеЗаполнено(НоваяСтрокаДанных.ИНН) Тогда
			НоваяСтрокаДанных.ИНН = "<" + НСтр("ru='Не заполнен'") + ">";
		КонецЕсли;
		
		Если НЕ ЗначениеЗаполнено(НоваяСтрокаДанных.СтраховойНомерПФР) Тогда
			НоваяСтрокаДанных.СтраховойНомерПФР = "<" + НСтр("ru='Не заполнен'") + ">";
		КонецЕсли;
		
		Если НЕ ЗначениеЗаполнено(НоваяСтрокаДанных.ДокументПредставление) Тогда
			НоваяСтрокаДанных.ДокументПредставление = "<" + НСтр("ru='Не заполнен'") + ">";
		КонецЕсли;
		
		ФИО = СокрЛП(НоваяСтрокаДанных.ФИОПолные);
		Пока СтрНайти(ФИО, "  ") > 0 Цикл
			ФИО = СтрЗаменить(ФИО, "  ", " ");
		КонецЦикла;
		
	КонецЦикла;
	
	Элементы.ДанныеФизическихЛиц.ТекущаяСтрока = ДанныеФизическихЛиц[0].ПолучитьИдентификатор();
	
	Если ДанныеФизическихЛиц.Количество() = 1 Тогда
		
		ДанныеФизическогоЛица = ДанныеФизическихЛиц[0];
		ЗаполнитьЗначенияСвойств(ЭтаФорма, ДанныеФизическогоЛица);
		ФизическоеЛицоСсылка = ДанныеФизическогоЛица.ФизическоеЛицо;
		Элементы.ГруппаСтраницы.ТекущаяСтраница = Элементы.СтраницаОдногоЧеловека;
		
	Иначе
		
		Элементы.ГруппаСтраницы.ТекущаяСтраница = Элементы.СтраницаНесколькихЧеловек;
		Элементы.ДаЭтоТотКтоМнеНужен.Заголовок = НСтр("ru='Отмеченный человек тот, кто мне нужен'");
		
	КонецЕсли;
	
КонецПроцедуры

#КонецОбласти

#Область ОбработчикиКомандФормы

&НаКлиенте
Процедура ДаЭтоТотКтоМнеНужен(Команда)
	
	ВыбратьФизическоеЛицоИЗакрытьФорму();
	
КонецПроцедуры

&НаКлиенте
Процедура ОткрытьДанныеФизическогоЛица(Команда)
	
	ТекДанные = Элементы.ДанныеФизическихЛиц.ТекущиеДанные;
	
	Если ТекДанные <> Неопределено Тогда
		ОткрытьЛичныеДанные(ТекДанные.ФизическоеЛицо);
	КонецЕсли;
	
КонецПроцедуры

&НаКлиенте
Процедура Отменить(Команда)
	
	Закрыть();
	
КонецПроцедуры

&НаКлиенте
Процедура НетЭтоДругойЧеловек(Команда)
	
	Закрыть();
	
КонецПроцедуры

&НаКлиенте
Процедура ВыбратьФизическоеЛицоКонтекстноеМеню(Команда)
	
	ТекДанные = Элементы.ДанныеФизическихЛиц.ТекущиеДанные;
	
	Если ТекДанные <> Неопределено Тогда
		Закрыть(ТекДанные.ФизическоеЛицо);
	КонецЕсли;
	
КонецПроцедуры

&НаКлиенте
Процедура ДанныеФизическихЛицПометкаПриИзменении(Элемент)
	
	ТекущиеДанныеСотрудника = Элементы.ДанныеФизическихЛиц.ТекущиеДанные;
	
	Если ТекущиеДанныеСотрудника.Пометка Тогда
		
		ПомеченныеСтроки = ДанныеФизическихЛиц.НайтиСтроки(Новый Структура("Пометка", Истина));
		Для Каждого СтрокаПомеченныеСтроки Из ПомеченныеСтроки Цикл
			
			Если СтрокаПомеченныеСтроки <> ТекущиеДанныеСотрудника Тогда
				СтрокаПомеченныеСтроки.Пометка = Ложь;
			КонецЕсли;
			
		КонецЦикла;
		
	КонецЕсли;
	
КонецПроцедуры

#КонецОбласти

#Область СлужебныеПроцедурыИФункции

&НаКлиенте
Процедура ОткрытьЛичныеДанные(ФизическоеЛицоСсылка)	
	
	ПараметрыОткрытияФормы = Новый Структура;
	ПараметрыОткрытияФормы.Вставить("Ключ", ФизическоеЛицоСсылка);
	ПараметрыОткрытияФормы.Вставить("ТолькоПросмотр", Истина);
	
	ОткрытьФорму("Справочник.ФизическиеЛица.ФормаОбъекта", ПараметрыОткрытияФормы, ЭтаФорма, , , , ,РежимОткрытияОкнаФормы.БлокироватьОкноВладельца);
	
КонецПроцедуры

&НаКлиенте
Процедура ВыбратьФизическоеЛицоИЗакрытьФорму()
	
	Если ДанныеФизическихЛиц.Количество() > 1 Тогда
		
		ПомеченныеСтроки = ДанныеФизическихЛиц.НайтиСтроки(Новый Структура("Пометка", Истина));
		Если ПомеченныеСтроки.Количество() > 0 Тогда
			Закрыть(ПомеченныеСтроки[0].ФизическоеЛицо);
		Иначе
			ПоказатьПредупреждение(, НСтр("ru = 'Необходимо установить отметку напротив нужного человека'"));
		КонецЕсли;
		
	Иначе
		Закрыть(ДанныеФизическихЛиц[0].ФизическоеЛицо);
	КонецЕсли;
	
КонецПроцедуры

#КонецОбласти