#Область ОбработчикиСобытий
Процедура ПриЗаписи(Отказ)
	
	УчетОстановочных = Константы.торо_ИспользоватьОстановочныеРемонты.Получить();
	Константы.торо_ИспользуютсяОстановочныеИСостояния.Установить(ЭтотОбъект.Значение И УчетОстановочных);
	
	УчетНаработки = Константы.торо_УчетНаработкиОборудования.Получить();
	Константы.торо_ИспользуютсяНаработкаИСостояния.Установить(ЭтотОбъект.Значение И УчетНаработки);
	
КонецПроцедуры
#КонецОбласти
