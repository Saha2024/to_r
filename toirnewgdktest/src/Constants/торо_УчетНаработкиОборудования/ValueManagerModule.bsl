#Область ОбработчикиСобытий

Процедура ПриЗаписи(Отказ)
	УчетСостояний = Константы.торо_УчетСостоянияОборудования.Получить();
	Константы.торо_ИспользуютсяНаработкаИСостояния.Установить(ЭтотОбъект.Значение И УчетСостояний);
КонецПроцедуры

#КонецОбласти