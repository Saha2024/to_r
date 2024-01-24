#Область ОбработчикиСобытийФормы

&НаСервере
Процедура ПриСозданииНаСервере(Отказ, СтандартнаяОбработка)
	
	Документ = Параметры.Документ;
	ДокументID = Параметры.ДокументID;
	ДокументТип = Параметры.ДокументТип;
	
	Если ЗначениеЗаполнено(ДокументID) Тогда
		
		ПрочитатьСостоянияДокумента(Параметры);
		
		Заголовок = СтроковыеФункцииКлиентСервер.ПодставитьПараметрыВСтроку(
			НСтр("ru = 'Состояние документа ""%1""'"),
			Документ);
		
	Иначе
		
		Если ДокументТип = "DMIncomingDocument" Тогда
			СостояниеРегистрация = НСтр("ru='На регистрации'");
			СостояниеРегистрацияID = "НаРегистрации"; //@NON-NLS-1
			СостояниеРегистрацияТип = "DMDocumentStatus";
		Иначе
			СостояниеРегистрация = НСтр("ru='Проект'");
			СостояниеРегистрацияID = "Проект"; //@NON-NLS-1
			СостояниеРегистрацияТип = "DMDocumentStatus";
		КонецЕсли;
		
	КонецЕсли;
	
	Если Параметры.ТолькоПросмотр Тогда
		
		ТолькоПросмотр = Истина;
		Элементы.ФормаЗаписать.Видимость = Ложь;
		Элементы.ФормаОтмена.Заголовок = НСтр("ru = 'Закрыть'");
		
	КонецЕсли;
	
	Элементы.СостояниеСогласование.СписокВыбора.Очистить();
	Элементы.СостояниеСогласование.СписокВыбора.Добавить("НаСогласовании", НСтр("ru='На согласовании'")); //@NON-NLS-1
	Элементы.СостояниеСогласование.СписокВыбора.Добавить("Согласован", НСтр("ru='Согласован'")); //@NON-NLS-1
	Элементы.СостояниеСогласование.СписокВыбора.Добавить("НеСогласован", НСтр("ru='Не согласован'")); //@NON-NLS-1
	
	Элементы.СостояниеУтверждение.СписокВыбора.Очистить();
	Элементы.СостояниеУтверждение.СписокВыбора.Добавить("НаУтверждении", НСтр("ru='На утверждении'")); //@NON-NLS-1
	Элементы.СостояниеУтверждение.СписокВыбора.Добавить("Утвержден", НСтр("ru='Утвержден'")); //@NON-NLS-1
	Элементы.СостояниеУтверждение.СписокВыбора.Добавить("НеУтвержден", НСтр("ru='Не утвержден'")); //@NON-NLS-1
	
	Элементы.СостояниеПодписание.СписокВыбора.Очистить();
	Элементы.СостояниеПодписание.СписокВыбора.Добавить("НаПодписании", НСтр("ru='На подписании'")); //@NON-NLS-1
	Элементы.СостояниеПодписание.СписокВыбора.Добавить("Подписан", НСтр("ru='Подписан'")); //@NON-NLS-1
	Элементы.СостояниеПодписание.СписокВыбора.Добавить("Отклонен", НСтр("ru='Отклонен'")); //@NON-NLS-1
	
	Элементы.СостояниеРегистрация.СписокВыбора.Очистить();
	Элементы.СостояниеРегистрация.СписокВыбора.Добавить("Проект", НСтр("ru='Проект'")); //@NON-NLS-1
	Элементы.СостояниеРегистрация.СписокВыбора.Добавить("НаРегистрации", НСтр("ru='На регистрации'")); //@NON-NLS-1
	Элементы.СостояниеРегистрация.СписокВыбора.Добавить("Зарегистрирован", НСтр("ru='Зарегистрирован'")); //@NON-NLS-1
	Элементы.СостояниеРегистрация.СписокВыбора.Добавить("НеЗарегистрирован", НСтр("ru='Не зарегистрирован'")); //@NON-NLS-1
	
	Элементы.СостояниеРассмотрение.СписокВыбора.Очистить();
	Элементы.СостояниеРассмотрение.СписокВыбора.Добавить("НаРассмотрении", НСтр("ru='На рассмотрении'")); //@NON-NLS-1
	Элементы.СостояниеРассмотрение.СписокВыбора.Добавить("Рассмотрен", НСтр("ru='Рассмотрен'")); //@NON-NLS-1
	
	Элементы.СостояниеИсполнение.СписокВыбора.Очистить();
	Элементы.СостояниеИсполнение.СписокВыбора.Добавить("НаИсполнении", НСтр("ru='На исполнении'")); //@NON-NLS-1
	Элементы.СостояниеИсполнение.СписокВыбора.Добавить("Исполнен", НСтр("ru='Исполнен'")); //@NON-NLS-1
	
	ОбновитьВидимость();
	
КонецПроцедуры

&НаКлиенте
Процедура ПриОткрытии(Отказ)
	
	#Если ВебКлиент Тогда
	Если Не Параметры.ТолькоПросмотр Тогда
	
	Элементы.СостояниеСогласование.РедактированиеТекста = Ложь;
	Элементы.СостояниеУтверждение.РедактированиеТекста = Ложь;
	Элементы.СостояниеПодписание.РедактированиеТекста = Ложь;
	Элементы.СостояниеРегистрация.РедактированиеТекста = Ложь;
	Элементы.СостояниеРассмотрение.РедактированиеТекста = Ложь;
	Элементы.СостояниеИсполнение.РедактированиеТекста = Ложь;
	
	Элементы.СостояниеСогласование.КнопкаОчистки = Истина;
	Элементы.СостояниеУтверждение.КнопкаОчистки = Истина;
	Элементы.СостояниеПодписание.КнопкаОчистки = Истина;
	Элементы.СостояниеРегистрация.КнопкаОчистки = Истина;
	Элементы.СостояниеРассмотрение.КнопкаОчистки = Истина;
	Элементы.СостояниеИсполнение.КнопкаОчистки = Истина;
	
	КонецЕсли;
	#КонецЕсли
	
КонецПроцедуры

#КонецОбласти

#Область ОбработчикиСобытийЭлементовШапкиФормы

&НаКлиенте
Процедура СостояниеСогласованиеПриИзменении(Элемент)
	
	ПриИзмененииСостояния(Элемент);
	
КонецПроцедуры

&НаКлиенте
Процедура СостояниеУтверждениеПриИзменении(Элемент)
	
	ПриИзмененииСостояния(Элемент);
	
КонецПроцедуры

&НаКлиенте
Процедура СостояниеПодписаниеПриИзменении(Элемент)
	
	ПриИзмененииСостояния(Элемент);
	
КонецПроцедуры

&НаКлиенте
Процедура СостояниеРегистрацияПриИзменении(Элемент)
	
	ПриИзмененииСостояния(Элемент);
	
КонецПроцедуры

&НаКлиенте
Процедура СостояниеРассмотрениеПриИзменении(Элемент)
	
	ПриИзмененииСостояния(Элемент);
	
КонецПроцедуры

&НаКлиенте
Процедура СостояниеИсполнениеПриИзменении(Элемент)
	
	ПриИзмененииСостояния(Элемент);
	
КонецПроцедуры

#КонецОбласти

#Область ОбработчикиКомандФормы

&НаКлиенте
Процедура ОК(Команда)
	
	ОчиститьСообщения();
	Если Не ПроверитьЗаполнение() Тогда
		Возврат;
	КонецЕсли;
	
	Результат = Новый Структура;
	
	Реквизиты = Новый Структура;
	Реквизиты.Вставить("СостояниеСогласование");
	Реквизиты.Вставить("СостояниеУтверждение");
	Реквизиты.Вставить("СостояниеПодписание");
	Реквизиты.Вставить("СостояниеРегистрация");
	Реквизиты.Вставить("СостояниеРассмотрение");
	Реквизиты.Вставить("СостояниеИсполнение");
	Для Каждого Реквизит Из Реквизиты Цикл
		ИмяРеквизита = Реквизит.Ключ;
		Результат.Вставить(ИмяРеквизита, ЭтотОбъект[ИмяРеквизита]);
		Результат.Вставить(ИмяРеквизита + "ID", ЭтотОбъект[ИмяРеквизита + "ID"]);
		Результат.Вставить(ИмяРеквизита + "Тип", ЭтотОбъект[ИмяРеквизита + "Тип"]);
	КонецЦикла;
	
	Закрыть(Результат);
	
КонецПроцедуры

&НаКлиенте
Процедура Отмена(Команда)
	
	Закрыть();
	
КонецПроцедуры

&НаСервере
Процедура ПрочитатьСостоянияДокумента(Параметры)
	
	Реквизиты = Новый Структура;
	Реквизиты.Вставить("СостояниеСогласование");
	Реквизиты.Вставить("СостояниеУтверждение");
	Реквизиты.Вставить("СостояниеРегистрация");
	Реквизиты.Вставить("СостояниеРассмотрение");
	Реквизиты.Вставить("СостояниеИсполнение");
	Если Параметры.Свойство("СостояниеПодписание") Тогда
		Реквизиты.Вставить("СостояниеПодписание");
	КонецЕсли;
	Для Каждого Реквизит Из Реквизиты Цикл
		ИмяРеквизита = Реквизит.Ключ;
		ЭтотОбъект[ИмяРеквизита] = Параметры[ИмяРеквизита];
		ЭтотОбъект[ИмяРеквизита + "ID"] = Параметры[ИмяРеквизита + "ID"];
		ЭтотОбъект[ИмяРеквизита + "Тип"] = Параметры[ИмяРеквизита + "Тип"];
	КонецЦикла;
	
КонецПроцедуры

#КонецОбласти

#Область СлужебныеПроцедурыИФункции

&НаКлиенте
Процедура ПриИзмененииСостояния(Элемент)
	
	ЭтотОбъект[Элемент.Имя] = Элемент.ТекстРедактирования;
	Если ЗначениеЗаполнено(ЭтотОбъект[Элемент.Имя]) Тогда
		ЭтотОбъект[Элемент.Имя + "Тип"] = "DMDocumentStatus";
	Иначе
		ЭтотОбъект[Элемент.Имя + "Тип"] = "";
	КонецЕсли;
	
КонецПроцедуры

&НаСервере
Процедура ОбновитьВидимость()
	
	ИспользуетсяПодписание = Ложь;
	
	Если ДокументТип = "DMInternalDocument" Тогда
		Прокси = ИнтеграцияС1СДокументооборотБазоваяФункциональностьПовтИсп.ПолучитьПрокси();
		ДокументXDTO = ИнтеграцияС1СДокументооборотБазоваяФункциональность.ПолучитьОбъект(
			Прокси,
			ДокументТип,
			ДокументID,
			"documentType");
		Если ИнтеграцияС1СДокументооборотБазоваяФункциональность.СвойствоУстановлено(
				ДокументXDTO.documentType, "useSigningByManager") Тогда
			ИспользуетсяПодписание = ДокументXDTO.documentType.useSigningByManager;
		КонецЕсли;
	КонецЕсли;
	
	Если ИспользуетсяПодписание Тогда
		Элементы.ГруппаУтверждение.Видимость = Ложь;
		Элементы.ГруппаПодписание.Видимость = Истина;
	Иначе
		Элементы.ГруппаУтверждение.Видимость = Истина;
		Элементы.ГруппаПодписание.Видимость = Ложь;
	КонецЕсли;
	
КонецПроцедуры

#КонецОбласти
