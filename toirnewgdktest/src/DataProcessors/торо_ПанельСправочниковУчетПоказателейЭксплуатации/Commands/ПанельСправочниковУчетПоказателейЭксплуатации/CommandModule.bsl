
#Область ОбработчикиСобытий

&НаКлиенте
Процедура ОбработкаКоманды(ПараметрКоманды, ПараметрыВыполненияКоманды)
	
	#Если ВебКлиент Тогда
	ОкноОткрытияПанели = ПараметрыВыполненияКоманды.Окно;
	#Иначе
	ОкноОткрытияПанели = ПараметрыВыполненияКоманды.Источник;
	#КонецЕсли
	
	ОткрытьФорму("Обработка.торо_ПанельСправочниковУчетПоказателейЭксплуатации.Форма.ПанельСправочниковУчетПоказателейЭксплуатации", , ПараметрыВыполненияКоманды.Источник, ПараметрыВыполненияКоманды.Уникальность, ОкноОткрытияПанели);
	
КонецПроцедуры

#КонецОбласти