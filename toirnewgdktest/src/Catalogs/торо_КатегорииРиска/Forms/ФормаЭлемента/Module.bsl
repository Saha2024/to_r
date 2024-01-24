#Область ОбработчикиСобытийФормы
&НаСервере
Процедура ПриСозданииНаСервере(Отказ, СтандартнаяОбработка)
	
	МультиязычностьСервер.ПриСозданииНаСервере(ЭтотОбъект, Объект);
	
	Если ЗначениеЗаполнено(Параметры.ЗначениеКопирования) Тогда
		ОтобразитьЦветНаФорме(Параметры.ЗначениеКопирования.ПолучитьОбъект());
	КонецЕсли;
	
	Элементы.Цвет.Доступность = Доступность И (ПравоДоступа("Изменение", Метаданные.Справочники.торо_КатегорииРиска)
											ИЛИ ПравоДоступа("Добавление", Метаданные.Справочники.торо_КатегорииРиска));
	
КонецПроцедуры

&НаКлиенте
Процедура ПриОткрытии(Отказ)
	ПроставитьАвтоотметкиНезаполненного();
КонецПроцедуры

&НаСервере
Процедура ПриЧтенииНаСервере(ТекущийОбъект)
	
	МультиязычностьСервер.ПриЧтенииНаСервере(ЭтотОбъект, ТекущийОбъект);
	ОтобразитьЦветНаФорме(ТекущийОбъект);
	
КонецПроцедуры

&НаСервере
Процедура ПередЗаписьюНаСервере(Отказ, ТекущийОбъект, ПараметрыЗаписи)
	
	МультиязычностьСервер.ПередЗаписьюНаСервере(ТекущийОбъект);
	
	ТекущийОбъект.Цвет = Новый ХранилищеЗначения(Цвет);
	
КонецПроцедуры

&НаСервере
Процедура ПослеЗаписиНаСервере(ТекущийОбъект, ПараметрыЗаписи)
	
	МультиязычностьСервер.ПриЧтенииНаСервере(ЭтотОбъект, ТекущийОбъект);
	
КонецПроцедуры

#КонецОбласти

#Область ОбработчикиСобытийЭлементовШапкиФормы

&НаКлиенте
Процедура Подключаемый_Открытие(Элемент, СтандартнаяОбработка)
	МультиязычностьКлиент.ПриОткрытии(ЭтотОбъект, Объект, Элемент, СтандартнаяОбработка);
КонецПроцедуры

&НаКлиенте
Процедура ДатаНачала_КоличествоПериодовПриИзменении(Элемент)
	ПроставитьАвтоотметкиНезаполненного();
КонецПроцедуры

&НаКлиенте
Процедура ДатаОкончания_КоличествоПериодовПриИзменении(Элемент)
	ПроставитьАвтоотметкиНезаполненного();
КонецПроцедуры

&НаКлиенте
Процедура ДатаПлановаяКрайняя_КоличествоПериодовПриИзменении(Элемент)
	ПроставитьАвтоотметкиНезаполненного();
КонецПроцедуры
 
#КонецОбласти

#Область СлужебныеПроцедурыИФункции

&НаКлиенте 
Процедура ПроставитьАвтоотметкиНезаполненного()
	
	Элементы.ДатаНачала_Период.АвтоОтметкаНезаполненного          = ЗначениеЗаполнено(Объект.ДатаНачала_КоличествоПериодов); 
	Элементы.ДатаОкончания_Период.АвтоОтметкаНезаполненного       = ЗначениеЗаполнено(Объект.ДатаОкончания_КоличествоПериодов); 
	Элементы.ДатаПлановаяКрайняя_Период.АвтоОтметкаНезаполненного = ЗначениеЗаполнено(Объект.ДатаПлановаяКрайняя_КоличествоПериодов); 
	
КонецПроцедуры

&НаСервере
Процедура ОтобразитьЦветНаФорме(ТекущийОбъект)
	
	ЦветИзХранилища = ТекущийОбъект.Цвет.Получить();
	Если Не ЦветИзХранилища = Неопределено Тогда
		Цвет = ЦветИзХранилища;
	КонецЕсли;
	
КонецПроцедуры

#КонецОбласти

  
