#Область ОбработчикиСобытийФормы

&НаСервере
Процедура проф_ПриСозданииНаСервереПосле(Отказ, СтандартнаяОбработка)
	
	//++ Проф-ИТ, #265, Башинская А.Ю., 25.09.2023
	НовыйЭлемент = Элементы.Добавить("ВидРемонта", Тип("ПолеФормы"));
	НовыйЭлемент.Вид = ВидПоляФормы.ПолеВвода;
	НовыйЭлемент.ПутьКДанным = "Объект.проф_ВидРемонта";      
	Элементы.Переместить(НовыйЭлемент, ЭтаФорма, Элементы.Комментарий);  
	НовыйЭлемент.АвтоМаксимальнаяШирина = Ложь;  
	НовыйЭлемент.РастягиватьПоГоризонтали = Истина;  
	//-- Проф-ИТ, #265, Башинская А.Ю., 25.09.2023
	
КонецПроцедуры

#КонецОбласти
