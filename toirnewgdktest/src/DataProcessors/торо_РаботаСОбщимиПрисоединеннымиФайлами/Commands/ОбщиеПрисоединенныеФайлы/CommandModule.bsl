#Область ОбработчикиСобытий

&НаКлиенте
Процедура ОбработкаКоманды(ПараметрКоманды, ПараметрыВыполненияКоманды)
	ПараметрыФормы = Новый Структура("Объект", ПараметрКоманды);
	ОткрытьФорму("Обработка.торо_РаботаСОбщимиПрисоединеннымиФайлами.Форма.СписокФайловПоОбъекту", ПараметрыФормы);
КонецПроцедуры

#КонецОбласти