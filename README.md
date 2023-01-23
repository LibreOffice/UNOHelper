# UNOHelper

UNOHelper offers helper functions for working with the LibreOffice UNO API.

It wraps commonly used functions from the UNO API for easier use.

## Build

The following applications have to be installed to compile UNOHelper:
* [JAVA JDK](http://www.oracle.com/technetwork/java/javase/downloads/index.html)
* [Apache Maven](https://maven.apache.org/download.cgi)
* [Git](http://git-scm.com/downloads/)

First fetch the sources:

```bash
git clone https://github.com/LibreOffice/UNOHelper.git
```

Then build the package:

```bash
mvn clean package
```

To create a single jar (to be used in non-Maven projects), run this command instead:

```bash
mvn clean assembly:single
```

## Release

First, make sure that Javadoc has no **errors** (warnings are ok though):

```bash
mvn javadoc:javadoc
```

Then do a dry run:

```bash
mvn -Dresume=false -DdryRun=true release:prepare
```

Now do the release. **Caution**: This creates two commits & a tag and pushes them directly to Git!)

```bash
mvn -Dresume=false release:prepare release:perform
```
