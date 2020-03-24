# UNOHelper

UNOHelper offers helper functions for working with the OpenOffice/LibreOffice 
UNO API. It wraps commonly used functions from the UNO API for easier use.

## Kompilieren

The following applications have to be installed to compile WollMux:
* [JAVA JDK](http://www.oracle.com/technetwork/java/javase/downloads/index.html)
* [Apache Maven](https://maven.apache.org/download.cgi)
* [Git](http://git-scm.com/downloads/)

Perform the following commands to download the sources and build the LibreOffice extension. Special dependencies of WollMux are hosted at
[Bintray](https://bintray.com/wollmux/WollMux), which is already configured as maven repository in pom.xml

```
git clone https://github.com/WollMux/UNOHelper.git
mvn clean package
```
