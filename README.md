# Проставление типа управляющей компании для адресов

## Задача
Имеются JSON файлы, где для каждой управляющей компании **(УК)** есть свой список обслуживаемых адресов. Необходимо сделать выгрузку адресов по этому региону, найти адреса, которые покрывают УК и проставить необходимую информацию (тип УК, номер лицензии, дата выдачи лицензии итд.) 

## Краткий алогритм работы
- Выгружаем из базы все адреса по этому региону
- Находим и вставляем дочерние
- Вставляем дополнительную информацию для каждого адреса (площадь, описание, владельцы)
- Обрабатываем JSON (В основном только парсим адреса) и записываем все в какой-нибудь словарь
- Проходимся по всем адресам, ищем похожие в словаре, проставляем инфу
- Делим на разные категории (по типу УК, по коду жилого помещения, по наличию УК итд.)

*Спать хочу потом еще попишу*
