/*-
 * #%L
 * UNOHelper
 * %%
 * Copyright (C) 2005 - 2023 The Document Foundation
 * %%
 * Licensed under the EUPL, Version 1.1 or – as soon they will be
 * approved by the European Commission - subsequent versions of the
 * EUPL (the "Licence");
 *
 * You may not use this work except in compliance with the Licence.
 * You may obtain a copy of the Licence at:
 *
 * http://ec.europa.eu/idabc/eupl5
 *
 * Unless required by applicable law or agreed to in writing, software
 * distributed under the Licence is distributed on an "AS IS" basis,
 * WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 * See the Licence for the specific language governing permissions and
 * limitations under the Licence.
 * #L%
 */
package org.libreoffice.ext.unohelper.common;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.Collections;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import com.sun.star.beans.PropertyVetoException;
import com.sun.star.beans.UnknownPropertyException;
import com.sun.star.beans.XPropertyChangeListener;
import com.sun.star.beans.XPropertySet;
import com.sun.star.beans.XPropertySetInfo;
import com.sun.star.beans.XVetoableChangeListener;
import com.sun.star.container.ElementExistException;
import com.sun.star.container.NoSuchElementException;
import com.sun.star.embed.ElementModes;
import com.sun.star.embed.InvalidStorageException;
import com.sun.star.embed.StorageWrappedTargetException;
import com.sun.star.embed.XEncryptionProtectedSource;
import com.sun.star.embed.XStorage;
import com.sun.star.embed.XTransactedObject;
import com.sun.star.embed.XTransactionBroadcaster;
import com.sun.star.embed.XTransactionListener;
import com.sun.star.io.IOException;
import com.sun.star.io.XInputStream;
import com.sun.star.io.XOutputStream;
import com.sun.star.io.XSeekable;
import com.sun.star.io.XStream;
import com.sun.star.lang.WrappedTargetException;
import com.sun.star.lang.XComponent;
import com.sun.star.lang.XEventListener;
import com.sun.star.packages.NoEncryptionException;
import com.sun.star.packages.WrongPasswordException;
import com.sun.star.uno.Type;

/**
 * Diese Klasse definiert die notwendigen Services, um mit mit der Methode
 * XStorageBasedDocument.storeToStorage(...) XML-Daten aktiver Dokumente in den
 * Hauptspeicher zu speichern. Dazu wird mit {@link #createByteArrayStorage()}
 * ein Objekt erzeugt, das die Daten im Hauptspeicher verwaltet und aus
 * Java-Sicht besonders leicht weiter verarbeitet werden kann. Mit
 * {@link #createXStorage(org.libreoffice.ext.unohelper.common.MemoryStorage.ByteArrayStorage)}
 * wird ein entsprechender UNO-Service erzeugt, der für storeToStorage(...)
 * geeignet ist.
 *
 * @author Christoph Lutz (D-III-ITD-D101)
 */
public class MemoryStorage
{

  /**
   * Der Logger, auf den Meldungen ausgegeben werden, wenn er gesetzt ist.
   */
  private static Logger logger = null;

  /**
   * Diese Methode erzeugt einen ByteArrayStorage in dem Dokumente unkomprimiert
   * und unverschlüsselt im Hauptspeicher abgelegt werden können.
   */
  public static ByteArrayStorage createByteArrayStorage()
  {
    return new ByteArrayStorage();
  }

  /**
   * Erzeugt eine Implementierung des Service com::sun::star::embed::Storage,
   * bei dem die Daten eines Dokuments in einem ByteArrayStorage abgelegt
   * werden.
   *
   * Achtung: Der UNO-Service definiert viele Methoden, die jedoch in der Praxis
   * nicht alle relevant sind. Um toten Code zu vermeiden, und da die
   * Spezifikation dieser Methoden in der UNO-API nicht aussagekräftig genug
   * ist, wurden die unnötigen Methoden daher nicht implementiert. Das birgt
   * aber auch das Risiko, dass neue OOo-Versionen in Zukunft das Storage anders
   * ansprechen und dabei nicht implementierte Methoden verwenden könnten. Als
   * Hilfestellung bei der Fehlersuche dient hier das public-field
   * {@link #logger}, in dem ein Logger hinterlegt werden kann und über den
   * verwendete, aber nicht implementierte Methoden erkannt werden können.
   */
  public static XStorage createXStorage(ByteArrayStorage bas)
  {
    return new StorageImpl(bas, "", ElementModes.WRITE);
  }

  /**
   * Logger-Interfaces für die Behandlung von Debug-Ausgaben dieser Klasse. Für
   * Debugging-Zwecke kann es sinnvoll und notwendig sein, einen eigenen Logger
   * zu definieren und diesen über setLogger(MemoryStorage.Logger) zu setzen.
   */
  public interface Logger
  {
    public void log(String s);
  }

  /**
   * Setzt den Logger auf den Meldungen dieser Klasse ausgegeben werden.
   *
   * @param newLogger
   */
  public static void setLogger(Logger newLogger)
  {
    logger = newLogger;
  }

  /**
   * Gibt eine Logger-Meldung mit der Nachricht s auf dem Logger aus, wenn
   * dieser != null ist.
   */
  private static void log(String s)
  {
    if (logger != null)
      logger.log(s);
  }

  /**
   * NotYetImplemented: Gibt eine Logger-Meldung mit der Nachricht
   * "NotYetImplemented: " + s für eine Methode aus, die derzeit nicht
   * implementiert ist.
   */
  private static void notYetImplemented(String s)
  {
    log("NotYetImplemented: " + s);
  }

  /**
   * Diese Klasse definiert ein Storage, das benannte ByteArray-Datenblöcke
   * aufnehmen kann. Das Storage enthält ausschließlich Elemente mit Daten und
   * hält diese in einer flachen Struktur ohne Hierarchie. Daher kennt es keine
   * Verzeichnisse, die einzelnen Elemente können aber Namen mit "/" enthalten.
   */
  public static class ByteArrayStorage
  {

    private HashMap<String, byte[]> mapNameToData;

    private HashMap<String, String> mapNameToMediaType;

    private String mediaType;

    private ByteArrayStorage()
    {
      this.mapNameToData = new HashMap<>();
      this.mapNameToMediaType = new HashMap<>();
      mediaType = null;
    }

    /**
     * Fügt das Element elementName mit den Daten data und dem MediaType
     * mediaType in dieses Storage ein.
     */
    public void setData(String elementName, byte[] data, String mediaType)
    {
      mapNameToData.put(elementName, data);
      mapNameToMediaType.put(elementName, mediaType);
    }

    /**
     * Liefert die Anzahl Datenbytes des Elements elementName zurück.
     */
    public int getSize(String elementName)
    {
      byte[] buffy = mapNameToData.get(elementName);
      if (buffy != null)
        return buffy.length;
      return 0;
    }

    /**
     * Liefert einen InputStream des Elements elementName zurück. Ist
     * elementName nicht definiert, so wird eine
     * java.util.NoSuchElementException geworfen.
     *
     * @throws java.util.NoSuchElementException
     */
    public InputStream getInputStream(String elementName)
    {
      byte[] data = mapNameToData.get(elementName);
      if (data != null)
        return new ByteArrayInputStream(data);
      else
        throw new java.util.NoSuchElementException(elementName);
    }

    /**
     * Liefert die namen aller Einträge des Storage in einer alphabetisch
     * sortierten Liste zurück.
     */
    public List<String> getElementNames()
    {
      List<String> list = new ArrayList<>(mapNameToData.keySet());
      Collections.sort(list);
      return list;
    }

    /**
     * Diese Methode erlaubt das Setzen des MediaTypes dieses Storages. Ist der
     * mediaType != null, so wird er beim Erzeugen des Manifests mittels
     * {@link #createManifest()} auch im Manifest manifestiert.
     *
     * @param mediaType
     *          Der mediaType für Textdokumente ist z.B.
     *          application/vnd.oasis.opendocument.text
     */
    public void setMediatype(String mediaType)
    {
      this.mediaType = mediaType;
    }

    /**
     * Diese Methode erzeugt eine Datei META-INF/manifest.xml, die die Elemente
     * des Storage beschreibt, und nimmt sie in das Storage mit auf.
     */
    public void createManifest()
    {
      StringBuilder manifest = new StringBuilder();
      manifest.append("<?xml version=\"1.0\" encoding=\"UTF-8\"?>\n");
      manifest.append(
          "<manifest:manifest xmlns:manifest=\"urn:oasis:names:tc:opendocument:xmlns:manifest:1.0\">\n");

      if (mediaType != null)
      {
        manifest.append("  <manifest:file-entry manifest:media-type=\"");
        manifest.append(mediaType);
        manifest.append("\" manifest:full-path=\"/\"/>\n");
      }

      for (Map.Entry<String, String> entry : mapNameToMediaType.entrySet())
      {
        manifest.append("  <manifest:file-entry manifest:media-type=\"");
        manifest.append(entry.getValue());
        manifest.append("\" manifest:full-path=\"");
        manifest.append(entry.getKey());
        manifest.append("\"/>\n");
      }

      manifest.append("</manifest:manifest>");
      mapNameToData.put("META-INF/manifest.xml",
          manifest.toString().getBytes());
    }
  }

  /**
   * Diese Klasse implementiert den UNO-Service com::sun::star::embed::Storage.
   * Viele Methoden müssen definiert sein, damit der Service Storage von der
   * UNO-Bridge korrekt anerkannt wird und ohne Fehler verwendet werden kann. In
   * der Praxis sind aber viele der definierten Methoden überflüssig. Sie werden
   * daher auch hier nicht implementiert, sonder nur definiert. Normalerweise
   * sollte dies keine Probleme bereiten. Sollte es damit doch einmal Probleme
   * geben, kann mit Hilfe des Loggers erkannt werden, ob eine nicht
   * implementierte Methode verwendet wurde und diese implementiert werden.
   *
   * @author Christoph Lutz (D-III-ITD-D101)
   */
  private static class StorageImpl extends SimplePropertySet
      implements XStorage, XTransactedObject, XTransactionBroadcaster,
      XEncryptionProtectedSource
  {

    private final String namePrefix;

    private final ByteArrayStorage bas;

    private boolean isRoot;

    private StorageImpl(ByteArrayStorage bas, String namePrefix, int openMode)
    {
      this.namePrefix = namePrefix;
      this.bas = bas;
      this.isRoot = namePrefix.length() == 0;
      props.put("OpenMode", openMode);
      props.put("URL", "mystore:///" + namePrefix);
      props.put("MediaType", "");
      props.put("MediaTypeFallbackIsUsed", Boolean.FALSE);
      props.put("IsRoot", isRoot);
      props.put("RepairPackage", Boolean.FALSE);
      props.put("HasEncryptedEntries", Boolean.FALSE);
    }

    @Override
    public XStream cloneEncryptedStreamElement(String arg0, String arg1)
        throws NoEncryptionException, WrongPasswordException, IOException,
        StorageWrappedTargetException
    {
      notYetImplemented(this + ".cloneEncryptedStreamElement" + arg0);
      return null;
    }

    @Override
    public XStream cloneStreamElement(String arg0)
        throws WrongPasswordException, IOException,
        StorageWrappedTargetException
    {
      notYetImplemented(this + ".cloneStreamElement " + arg0);
      return null;
    }

    @Override
    public void copyElementTo(String arg0, XStorage arg1, String arg2)
        throws NoSuchElementException, ElementExistException, IOException,
        StorageWrappedTargetException
    {
      notYetImplemented(this + ".copyElementTo " + arg0);
    }

    @Override
    public void copyLastCommitTo(XStorage arg0)
        throws IOException, StorageWrappedTargetException
    {
      notYetImplemented(this + ".copyLastCommitTo " + arg0);
    }

    @Override
    public void copyStorageElementLastCommitTo(String arg0, XStorage arg1)
        throws IOException, StorageWrappedTargetException
    {
      notYetImplemented(this + ".copyStorageElementLastCommitTo " + arg0);
    }

    @Override
    public void copyToStorage(XStorage arg0)
        throws IOException, StorageWrappedTargetException
    {
      notYetImplemented(this + ".copyToStorage " + arg0);
    }

    @Override
    public boolean isStorageElement(String arg0)
        throws NoSuchElementException, InvalidStorageException
    {
      notYetImplemented(this + ".isStorageElement " + arg0);
      return false;
    }

    @Override
    public boolean isStreamElement(String arg0)
        throws NoSuchElementException, InvalidStorageException
    {
      notYetImplemented(this + ".isStreamElement " + arg0);
      return false;
    }

    @Override
    public void moveElementTo(String arg0, XStorage arg1, String arg2)
        throws NoSuchElementException, ElementExistException, IOException,
        StorageWrappedTargetException
    {
      notYetImplemented(this + ".moveElementTo " + arg0);
    }

    @Override
    public XStream openEncryptedStreamElement(String arg0, int arg1,
        String arg2) throws NoEncryptionException, WrongPasswordException,
        IOException, StorageWrappedTargetException
    {
      notYetImplemented(this + ".openEncryptedStreamElement " + arg0);
      return null;
    }

    @Override
    public XStorage openStorageElement(String arg0, int arg1)
        throws IOException, StorageWrappedTargetException
    {

      log(this + ".openStorageElement " + arg0);
      return new StorageImpl(bas, namePrefix + arg0 + "/", arg1);
    }

    @Override
    public XStream openStreamElement(String name, int openMode)
        throws WrongPasswordException, IOException,
        StorageWrappedTargetException
    {

      log(this + ".openStreamElement " + name);
      return new StreamElementImpl(bas, namePrefix + name, openMode);
    }

    @Override
    public void removeElement(String arg0) throws NoSuchElementException,
        IOException, StorageWrappedTargetException
    {
      notYetImplemented(this + ".removeElement " + arg0);
    }

    @Override
    public void renameElement(String arg0, String arg1)
        throws NoSuchElementException, ElementExistException, IOException,
        StorageWrappedTargetException
    {
      notYetImplemented(this + ".renameElement " + arg0);
    }

    @Override
    public Object getByName(String arg0)
        throws NoSuchElementException, WrappedTargetException
    {
      notYetImplemented(this + ".getByName " + arg0);
      return null;
    }

    @Override
    public String[] getElementNames()
    {
      notYetImplemented(this + ".getElementNames");
      return new String[] {};
    }

    @Override
    public boolean hasByName(String arg0)
    {
      log(this + ".hasByName " + arg0);
      for (String name : bas.getElementNames())
      {
        name += "/";
        if (name.startsWith(namePrefix + arg0 + "/"))
          return true;
      }
      return false;
    }

    @Override
    public Type getElementType()
    {
      notYetImplemented(this + ".getElementType");
      return null;
    }

    @Override
    public boolean hasElements()
    {
      notYetImplemented(this + ".hasElements");
      return false;
    }

    @Override
    public void addEventListener(XEventListener arg0)
    {
      notYetImplemented(this + ".addEventListener");
    }

    @Override
    public void dispose()
    {
      notYetImplemented(this + ".dispose");
    }

    @Override
    public void removeEventListener(XEventListener arg0)
    {
      notYetImplemented(this + ".removeEventListener");
    }

    @Override
    public void commit() throws IOException, WrappedTargetException
    {
      log(this + ".commit");
      if (isRoot)
      {

        Object mediaType = props.get("MediaType");
        if (mediaType != null)
          bas.setMediatype(mediaType.toString());

        bas.createManifest();
      }
    }

    @Override
    public void revert() throws IOException, WrappedTargetException
    {
      notYetImplemented(this + ".revert");
    }

    @Override
    public void addTransactionListener(XTransactionListener arg0)
    {
      notYetImplemented(this + ".addTransactionListener");
    }

    @Override
    public void removeTransactionListener(XTransactionListener arg0)
    {
      notYetImplemented(this + ".removeTransactionListener");
    }

    @Override
    public void removeEncryption() throws IOException
    {
      notYetImplemented(this + ".removeEncryption");
    }

    @Override
    public void setEncryptionPassword(String arg0) throws IOException
    {
      notYetImplemented(this + ".setEncryptionPassword");
    }

    @Override
    public String toString()
    {
      return props.get("URL").toString();
    }
  }

  /**
   * Diese Klasse implementiert den UNO-Service
   * com::sun::star::embed::StorageStream. Viele Methoden müssen definiert sein,
   * damit der Service StorageStream von der UNO-Bridge korrekt anerkannt wird
   * und ohne Fehler verwendet werden kann. In der Praxis sind aber einige der
   * definierten Methoden überflüssig. Sie werden daher auch nicht
   * implementiert, sonder nur definiert. Normalerweise sollte dies keine
   * Probleme bereiten. Sollte es damit doch einmal Probleme geben, kann mit
   * Hilfe des Loggers erkannt werden, ob eine nicht implementierte Methode
   * verwendet wurde und diese implementiert werden.
   *
   * @author Christoph Lutz (D-III-ITD-D101)
   */
  private static class StreamElementImpl extends SimplePropertySet implements
      XStream, XPropertySet, XSeekable, XEncryptionProtectedSource, XComponent
  {

    private final String name;

    private final ByteArrayStorage bas;

    private final String cn;

    private final ByteArrayOutputStream baos = new ByteArrayOutputStream();

    private final XOutputStream xOutputStream = new XOutputStream()
    {
      @Override
      public void writeBytes(byte[] arg0) throws IOException
      {
        log(cn + ".XOutputStream.writeBytes (length=" + arg0.length + ")");
        try
        {
          baos.write(arg0);
        } catch (java.io.IOException e)
        {
          throw new IOException(e.toString());
        }
      }

      @Override
      public void flush() throws IOException
      {
        log(cn + ".XOutputStream.flush");
        flushIntern();
      }

      private void flushIntern()
      {
        bas.setData(name, baos.toByteArray(), getMediaType());
        props.put("Size", bas.getSize(name));
      }

      @Override
      public void closeOutput() throws IOException
      {
        log(cn + ".XOutputStream.closeOutput");
        flushIntern();
        try
        {
          baos.close();
        } catch (java.io.IOException e)
        {
          throw new IOException(e.toString());
        }
      }

      private String getMediaType()
      {
        Object o = props.get("MediaType");
        return (o == null) ? "" : o.toString();
      }
    };

    private final XInputStream xInputStream = new XInputStream()
    {
      @Override
      public void skipBytes(int arg0) throws IOException
      {
        notYetImplemented(cn + ".XInputStream.skipBytes");
      }

      @Override
      public int readSomeBytes(byte[][] arg0, int arg1) throws IOException
      {
        notYetImplemented(cn + ".XInputStream.readSomeBytes");
        return 0;
      }

      @Override
      public int readBytes(byte[][] arg0, int arg1) throws  IOException
      {
        notYetImplemented(cn + ".XInputStream.readBytes");
        return 0;
      }

      @Override
      public void closeInput() throws IOException
      {
        notYetImplemented(cn + ".XInputStream.closeInput");
      }

      @Override
      public int available() throws IOException
      {
        notYetImplemented(cn + ".XInputStream.available");
        return 0;
      }
    };

    public StreamElementImpl(ByteArrayStorage bas, String name, int openMode)
    {
      this.name = name;
      this.bas = bas;
      props.put("OpenMode", openMode);
      props.put("MediaType", "");
      props.put("IsCompressed", Boolean.FALSE);
      props.put("IsEncrypted", Boolean.FALSE);
      props.put("UseCommonStoragePasswordEncryption", Boolean.FALSE);
      props.put("RepairPackage", Boolean.FALSE);
      props.put("Size", 0);
      this.cn = toString();
    }

    @Override
    public XInputStream getInputStream()
    {
      log(this + ".getInputStream");
      return xInputStream;
    }

    @Override
    public XOutputStream getOutputStream()
    {
      log(this + ".getOutputStream");
      return xOutputStream;
    }

    @Override
    public String toString()
    {
      return "streamelement:///" + name;
    }

    @Override
    public long getLength() throws IOException
    {
      notYetImplemented(this + ".getLength");
      return 0;
    }

    @Override
    public long getPosition() throws IOException
    {
      notYetImplemented(this + ".getPosition");
      return 0;
    }

    @Override
    public void seek(long arg0) throws IOException
    {
      notYetImplemented(this + ".seek");
    }

    @Override
    public void removeEncryption() throws IOException
    {
      notYetImplemented(this + ".removeEncryption");
    }

    @Override
    public void setEncryptionPassword(String arg0) throws IOException
    {
      notYetImplemented(this + ".setEncryptionPassword");
    }

    @Override
    public void addEventListener(XEventListener arg0)
    {
      notYetImplemented(this + ".addEventListener");
    }

    @Override
    public void dispose()
    {
      notYetImplemented(this + ".dispose");
    }

    @Override
    public void removeEventListener(XEventListener arg0)
    {
      notYetImplemented(this + ".removeEventListener");
    }

  }

  /**
   * Diese Klasse implementiert ein einfaches PropertySet basierend auf einer
   * HashMap. Auch hier sind viele in der Praxis nicht benötigte Methoden nicht
   * implementiert, können aber bei Bedarf implementiert werden.
   *
   * @author Christoph Lutz (D-III-ITD-D101)
   */
  private static class SimplePropertySet implements XPropertySet
  {
    protected final HashMap<String, Object> props;

    public SimplePropertySet()
    {
      props = new HashMap<>();
    }

    @Override
    public void addPropertyChangeListener(String arg0,
        XPropertyChangeListener arg1)
        throws UnknownPropertyException, WrappedTargetException
    {
      notYetImplemented(this + ".addPropertyChangeListener");
    }

    @Override
    public void addVetoableChangeListener(String arg0,
        XVetoableChangeListener arg1)
        throws UnknownPropertyException, WrappedTargetException
    {
      notYetImplemented(this + ".addVetoableChangeListener");
    }

    @Override
    public XPropertySetInfo getPropertySetInfo()
    {
      notYetImplemented(this + ".getPropertySetInfo");
      return null;
    }

    @Override
    public Object getPropertyValue(String arg0)
        throws UnknownPropertyException, WrappedTargetException
    {

      log(this + ".getPropertyValue " + arg0);
      Object o = props.get(arg0);
      if (o != null)
        return o;
      else
        throw new UnknownPropertyException(arg0);
    }

    @Override
    public void removePropertyChangeListener(String arg0,
        XPropertyChangeListener arg1)
        throws UnknownPropertyException, WrappedTargetException
    {
      notYetImplemented(this + ".removePropertyChangeListener " + arg0);
    }

    @Override
    public void removeVetoableChangeListener(String arg0,
        XVetoableChangeListener arg1)
        throws UnknownPropertyException, WrappedTargetException
    {
      notYetImplemented(this + ".removeVetoableChangeListener " + arg0);
    }

    @Override
    public void setPropertyValue(String arg0, Object arg1)
        throws UnknownPropertyException, PropertyVetoException,
        WrappedTargetException
    {

      log(this + ".setPropertyValue " + arg0 + " " + arg1);
      props.put(arg0, arg1);
    }
  }
}
