Сервис управления MS Office
Порт указывается в .env файле
Доступны функции:
Конвертация в PDF/A
POST: {IP}:{PORT}/convert
form-data (key: file)
Список форматов:
word -  "doc", "docx", "docm", "rtf", "xml", "pdf", "odt", "txt", "wbk" 
powerPoint -"pptx", "pptm", "ppt" 
picture -  "jpg", "jpeg", "png", "tiff", "tif"
excel - "xls", "xlsx", "csv"
Форматы которых нет в списке попытаются конвертироваться через ворд