## UNOHelper

UNOHelper offers helper functions for working with the OpenOffice/LibreOffice 
UNO API. It wraps commonly used functions from the UNO API for easier use.

UNOHelper stellt Hilfsfunktionen für die Benutzung der UNO API von 
OpenOffice/LibreOffice. Häufig benötigte Funktionaltäten der UNO API werden
durch Wrapper-Funktionen zur Verfügung gestellt, um die Arbeit mit UNO zu
vereinfachen.

### Kompilieren

Zum kompilieren von UNOHelper wird jetzt [Maven](https://maven.apache.org) 
verwendet. Als Voraussetzung muss daher zunächst Maven installiert werden.

Einige Abhängigkeiten die zum Kompilieren benötigt werden, stehen in einem
eigenen Maven-Repository zur Verfügung. Dieses Repository muss in der Datei
.m2/settings.xml eingetragen werden.

```
<repository>
	<snapshots>
		<enabled>false</enabled>
	</snapshots>
	<id>bintray-eymux-WollMux</id>
	<name>bintray</name>
	<url>http://dl.bintray.com/eymux/WollMux</url>
</repository>

```

Danach kann UNOHelper mit den Standardbefehlen von Maven gebaut werden.

> mvn compile

Kompiliert die Sourcen.

> mvn package

Erzeugt eine JAR-Datei im Verzeichnis `target`.

> mvn install

Installiert die JAR-Datei im lokalen Maven-Repository. 

