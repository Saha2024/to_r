
&НаКлиенте
Процедура ОбработкаКоманды(ПараметрКоманды, ПараметрыВыполненияКоманды)
	
	ПараметрыФормы = Новый Структура("", );
	ОткрытьФорму("Справочник.проф_НастройкиСистемы.ФормаСписка", ПараметрыФормы, ПараметрыВыполненияКоманды.Источник,
		ПараметрыВыполненияКоманды.Уникальность, ПараметрыВыполненияКоманды.Окно,
		ПараметрыВыполненияКоманды.НавигационнаяСсылка);
		
КонецПроцедуры
