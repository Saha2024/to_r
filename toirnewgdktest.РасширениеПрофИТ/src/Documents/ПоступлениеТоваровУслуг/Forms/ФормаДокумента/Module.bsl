#Область ОбработчикиСобытийФормы

&НаСервере
Процедура проф_ПриСозданииНаСервереПосле(Отказ, СтандартнаяОбработка)
	
	//++ Проф-ИТ, #27, Соловьев А.А., 25.08.2023
	НовыйЭлемент  = Элементы.Добавить("ТоварыНазначение", Тип("ПолеФормы"), Элементы.Товары);
	НовыйЭлемент.Вид = ВидПоляФормы.ПолеВвода;      
	НовыйЭлемент.ПутьКДанным = "Объект.Товары.проф_Назначение";
	НовыйЭлемент.ТолькоПросмотр = Истина;
	Элементы.Переместить(НовыйЭлемент, Элементы.Товары, Элементы.ТоварыУпаковкаЕдиницыИзмерения);
	//-- Проф-ИТ, #27, Соловьев А.А., 25.08.2023
	
КонецПроцедуры

#КонецОбласти