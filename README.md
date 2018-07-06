# FalschparkmelderFFM
Auf Twitter verfolge ich seit Längerem die Diskussionen vieler radfahrender User mit der [Stadt Frankfurt/Main](https://twitter.com/stadt_ffm?lang=de) bzgl. der Großzahl an Falschparkern, die die bereits wenigen Radwege versperren. Durch einen Tweet von Twitter-User [FFMByBicycle](https://twitter.com/ffmbybicycle) bin ich darauf aufmerksam geworden, dass Falschparker mittlerweile auch per Email beim Ordnungsamt der Stadt Frankfurt angezeigt werden können. Er hostet hierfür die Info-Seite https://www.falschparken-frankfurt.info/

Um die tagtäglich gesammelten Vorfälle möglichst einfach und schnell dem Ordnungsamt melden zu können, habe ich ein schlankes Tool programmiert: den FalschparkmelderFFM.

Vorfälle lassen sich schnell eingeben und versenden. Bereits versendete Vorfälle sind sauber in einer Datenbank gesammelt und können analysiert werden. Die Koordinaten der gemeldeten Vorfälle lässt sich zudem exportieren und bspw. auf Google Maps oder [GPSVisualizer](http://www.gpsvisualizer.com) darstellen. So bekommt man eine gute Vorstellung, wo die Hotspots der Falschparker vorliegen und wo Radfahren definitiv keinen Spaß macht.


## Voraussetzungen
 - Windows 7/8/10
 - PowerShell
 - PowerShell Modul PSSQLite (https://github.com/RamblingCookieMonster/PSSQLite)
    - den Ordner **PSSQLite** in den Ordner **%homepath%\Documents\WindowsPowerShell\modules** kopieren


## Datenbank
Die Datenbank wird im SQLite Format gespeichert und lässt sich daher auch ohne FalschparkmelderFFM auslesen/bearbeiten. Hierfür empfehle ich den kostenlosen SQLite DB Browser (https://sqlitebrowser.org/).
