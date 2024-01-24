#Область ОбработчикиСобытийФормы

//++ Проф-ИТ, #168, Лавриненко Т.В.,21.08.2023
&НаСервере
Процедура проф_ПриСозданииНаСервереПосле(Отказ, СтандартнаяОбработка)
	ЭтаФорма.Список.ТекстЗапроса = СтрЗаменить(ЭтаФорма.Список.ТекстЗапроса,"КАК ВидыРемонтаСтрокой",
										"КАК ВидыРемонтаСтрокой, Документторо_ЗаявкаНаРемонт.проф_ТипЗаказа КАК ТипЗаказа, Документторо_ЗаявкаНаРемонт.проф_СтатусСогласования КАК проф_СтатусСогласования"); 
	
	НоваяКолонка = Элементы.Вставить("ТипЗаказа",
									Тип("ПолеФормы"),
									Элементы.Список,Элементы.Организация);
    НоваяКолонка.ПутьКДанным = "Список.ТипЗаказа"; 
	
	//++ Проф-ИТ, #238, Башинская А.Ю.,26.08.2023 
	НоваяКолонка = Элементы.Вставить("СтатусСогласования",
									Тип("ПолеФормы"),
									Элементы.Список);
    НоваяКолонка.ПутьКДанным = "Список.проф_СтатусСогласования";     
	НоваяКолонка.Заголовок = "Статус согласования";    
    //-- Проф-ИТ, #238, Башинская А.Ю.,26.08.2023
КонецПроцедуры
 //-- Проф-ИТ, #168, Лавриненко Т.В.,21.08.2023
 #КонецОбласти