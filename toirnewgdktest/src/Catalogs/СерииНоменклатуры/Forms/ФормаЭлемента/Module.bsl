
#Область ОбработчикиСобытийФормы

&НаСервере
Процедура ПриСозданииНаСервере(Отказ, СтандартнаяОбработка)
	
	// СтандартныеПодсистемы.Свойства
	ДополнительныеПараметры = Новый Структура;
	ДополнительныеПараметры.Вставить("ИмяЭлементаДляРазмещения", "ГруппаДополнительныеРеквизиты");
	УправлениеСвойствами.ПриСозданииНаСервере(ЭтотОбъект, ДополнительныеПараметры);
	// Конец СтандартныеПодсистемы.Свойства
	
	Если Параметры.Свойство("ВидНоменклатуры") Тогда
		Объект.ВидНоменклатуры = Параметры.ВидНоменклатуры;
	КонецЕсли;
	
	Если ЗначениеЗаполнено(Объект.ВидНоменклатуры) Тогда
	    Элементы.ВидНоменклатуры.ТолькоПросмотр = Истина;
	КонецЕсли; 
	
КонецПроцедуры

&НаКлиенте
Процедура ПриОткрытии(Отказ)
	
	// СтандартныеПодсистемы.Свойства
	УправлениеСвойствамиКлиент.ПослеЗагрузкиДополнительныхРеквизитов(ЭтотОбъект);
	// Конец СтандартныеПодсистемы.Свойства
	
	УправлениеФормой();
	
КонецПроцедуры

&НаСервере
Процедура ПриЧтенииНаСервере(ТекущийОбъект)
	
	// СтандартныеПодсистемы.Свойства
	УправлениеСвойствами.ПриЧтенииНаСервере(ЭтотОбъект, ТекущийОбъект);
	// Конец СтандартныеПодсистемы.Свойства

КонецПроцедуры

&НаКлиенте
Процедура ОбработкаОповещения(ИмяСобытия, Параметр, Источник)
	
	// СтандартныеПодсистемы.Свойства 
	Если УправлениеСвойствамиКлиент.ОбрабатыватьОповещения(ЭтотОбъект, ИмяСобытия, Параметр) Тогда
		ОбновитьЭлементыДополнительныхРеквизитов();
		УправлениеСвойствамиКлиент.ПослеЗагрузкиДополнительныхРеквизитов(ЭтотОбъект);
	КонецЕсли;
	// Конец СтандартныеПодсистемы.Свойства

КонецПроцедуры

&НаСервере
Процедура ОбработкаПроверкиЗаполненияНаСервере(Отказ, ПроверяемыеРеквизиты)
	
	// СтандартныеПодсистемы.Свойства
	УправлениеСвойствами.ОбработкаПроверкиЗаполнения(ЭтотОбъект, Отказ, ПроверяемыеРеквизиты);
	// Конец СтандартныеПодсистемы.Свойства

КонецПроцедуры

&НаСервере
Процедура ПередЗаписьюНаСервере(Отказ, ТекущийОбъект, ПараметрыЗаписи)
	
	// СтандартныеПодсистемы.Свойства
	УправлениеСвойствами.ПередЗаписьюНаСервере(ЭтотОбъект, ТекущийОбъект);
	// Конец СтандартныеПодсистемы.Свойства

КонецПроцедуры

#КонецОбласти

#Область ОбработчикиСобытийЭлементовШапкиФормы

&НаКлиенте
Процедура ВидНоменклатурыПриИзменении(Элемент)
	
	УправлениеФормой();
	
КонецПроцедуры

&НаКлиенте
Процедура ВидНоменклатурыНачалоВыбора(Элемент, ДанныеВыбора, СтандартнаяОбработка)
	СтандартнаяОбработка = Ложь;     
	ФормаВыбора = ПолучитьФорму("Справочник.ВидыНоменклатуры.ФормаВыбора" ,,Элемент);
	ЭлементОтбора = ФормаВыбора.Список.Отбор.Элементы.Добавить(Тип("ЭлементОтбораКомпоновкиДанных"));
	ЭлементОтбора.ВидСравнения = ВидСравненияКомпоновкиДанных.Равно;
	ЭлементОтбора.ЛевоеЗначение = Новый ПолеКомпоновкиДанных("ИспользоватьСерии");
	ЭлементОтбора.ПравоеЗначение = Истина;
	ЭлементОтбора.Использование = Истина;
	ФормаВыбора.Открыть();
КонецПроцедуры

&НаКлиенте
Процедура ВидНоменклатурыАвтоПодбор(Элемент, Текст, ДанныеВыбора, ПараметрыПолученияДанных, Ожидание, СтандартнаяОбработка)
	СтандартнаяОбработка = Ложь;
	ПараметрыПолученияДанных = Новый Структура("Отбор", Новый Структура("СтрокаПоиска, ИспользоватьСерии", Текст, Истина));
	ДанныеВыбора = ПолучитьДанныеВыбора(Тип("СправочникСсылка.ВидыНоменклатуры"), ПараметрыПолученияДанных);
КонецПроцедуры

&НаКлиенте
Процедура ВидНоменклатурыОкончаниеВводаТекста(Элемент, Текст, ДанныеВыбора, ПараметрыПолученияДанных, СтандартнаяОбработка)
	СтандартнаяОбработка = Ложь;
	ПараметрыПолученияДанных = Новый Структура("Отбор", Новый Структура("СтрокаПоиска, ИспользоватьСерии", Текст, Истина));
	ДанныеВыбора = ПолучитьДанныеВыбора(Тип("СправочникСсылка.ВидыНоменклатуры"), ПараметрыПолученияДанных);
КонецПроцедуры

#КонецОбласти

#Область ОбработчикиКомандФормы

// СтандартныеПодсистемы.Свойства
&НаКлиенте
Процедура Подключаемый_СвойстваВыполнитьКоманду(ЭлементИлиКоманда, НавигационнаяСсылка = Неопределено, СтандартнаяОбработка = Неопределено)
	УправлениеСвойствамиКлиент.ВыполнитьКоманду(ЭтотОбъект, ЭлементИлиКоманда, СтандартнаяОбработка);
КонецПроцедуры
// Конец СтандартныеПодсистемы.Свойства

#КонецОбласти

#Область СлужебныеПроцедурыИФункции

// СтандартныеПодсистемы.Свойства 
&НаСервере
Процедура ОбновитьЭлементыДополнительныхРеквизитов()
	УправлениеСвойствами.ОбновитьЭлементыДополнительныхРеквизитов(ЭтотОбъект);
КонецПроцедуры

&НаКлиенте
Процедура ОбновитьЗависимостиДополнительныхРеквизитов()
	УправлениеСвойствамиКлиент.ОбновитьЗависимостиДополнительныхРеквизитов(ЭтотОбъект);
КонецПроцедуры

&НаКлиенте
Процедура Подключаемый_ПриИзмененииДополнительногоРеквизита(Элемент)
	УправлениеСвойствамиКлиент.ОбновитьЗависимостиДополнительныхРеквизитов(ЭтотОбъект);
КонецПроцедуры
// Конец СтандартныеПодсистемы.Свойства

&НаКлиенте
Процедура УправлениеФормой()
	
	НастройкаИспользованияСерий = торо_ОбщегоНазначенияВызовСервера.ЗначениеРеквизитаОбъекта(Объект.ВидНоменклатуры, "НастройкаИспользованияСерий");
	
	Элементы.ГоденДо.Видимость = НастройкаИспользованияСерий = ПредопределенноеЗначение("Перечисление.НастройкиИспользованияСерийНоменклатуры.ПартияТоваровПоНомеруИСрокуГодности")
										ИЛИ НастройкаИспользованияСерий = ПредопределенноеЗначение("Перечисление.НастройкиИспользованияСерийНоменклатуры.ПартияТоваровПоСрокуГодности");

	Элементы.Номер.Видимость = НастройкаИспользованияСерий = ПредопределенноеЗначение("Перечисление.НастройкиИспользованияСерийНоменклатуры.ПартияТоваровПоНомеруИСрокуГодности")
										ИЛИ НастройкаИспользованияСерий = ПредопределенноеЗначение("Перечисление.НастройкиИспользованияСерийНоменклатуры.ПартияТоваровПоНомеру")
										ИЛИ НастройкаИспользованияСерий = ПредопределенноеЗначение("Перечисление.НастройкиИспользованияСерийНоменклатуры.ЭкземплярТовара");
	
КонецПроцедуры

#КонецОбласти
