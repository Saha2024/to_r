
#Область СлужебныйПрограммныйИнтерфейс

// См. СтандартныеПодсистемыКлиент.ПриПолученииСерверногоОповещения
Процедура ПриПолученииСерверногоОповещения(ИмяОповещения, Результат) Экспорт

	#Если Не МобильныйКлиент Тогда
		Если ИмяОповещения = торо_РаботаСУведомлениямиКлиентСервер.ИмяСерверногоОповещения() 
			И ОбщегоНазначенияВызовСервера.ХранилищеОбщихНастроекЗагрузить("НастройкиТОиР", "ПолучатьУведомленияВВидеВсплывающихПодсказок", Истина) Тогда
			
			МассивУведомлений = Результат;
			
			Для Каждого Уведомление из МассивУведомлений Цикл
				ПоказатьОповещениеПользователя(Уведомление.ТемаПисьма, ?(ЗначениеЗаполнено(Уведомление.Источник), ПолучитьНавигационнуюСсылку(Уведомление.Источник), ""), Уведомление.ТелоПисьма);
			КонецЦикла;
			
		КонецЕсли;
	#КонецЕсли

КонецПроцедуры

#КонецОбласти
