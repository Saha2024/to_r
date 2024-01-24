
#Область ОбработчикиСобытий

Процедура ОбработкаПолученияПолейПредставления(Поля, СтандартнаяОбработка)
	
	СтандартнаяОбработка = Ложь;
	Поля.Добавить("Наименование");
	Поля.Добавить("Владелец");
	
	МультиязычностьКлиентСервер.ОбработкаПолученияПолейПредставления(Поля, СтандартнаяОбработка);
	
КонецПроцедуры

Процедура ОбработкаПолученияПредставления(Данные, Представление, СтандартнаяОбработка)
	
	МультиязычностьКлиентСервер.ОбработкаПолученияПредставления(Данные, Представление, СтандартнаяОбработка);
	
	СтандартнаяОбработка = Ложь;
	Представление = Строка(Данные.Владелец)+" / "+?(ЗначениеЗаполнено(Представление), Представление, Данные.Наименование);
	
КонецПроцедуры

#Если Сервер Или ТолстыйКлиентОбычноеПриложение Или ВнешнееСоединение Тогда

Процедура ОбработкаПолученияДанныхВыбора(ДанныеВыбора, Параметры, СтандартнаяОбработка)
	МультиязычностьСервер.ОбработкаПолученияДанныхВыбора(ДанныеВыбора, Параметры, СтандартнаяОбработка, Метаданные.Справочники.торо_ВариантыАнализаПоказателейKPI);
КонецПроцедуры

#КонецЕсли

#КонецОбласти