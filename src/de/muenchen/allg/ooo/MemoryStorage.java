/*
 * Dateiname: MemoryStorage.java
 * Projekt  : UNOHelper
 * Funktion : Definiert einen UNO-Service Storage zum Speichern von
 *            Dokumenten in den Hauptspeicher. 
 * 
 * Copyright (c) 2008 Landeshauptstadt München
 *
 * This program is free software: you can redistribute it and/or modify
 * it under the terms of the European Union Public Licence (EUPL),
 * version 1.0.
 *
 * This program is distributed in the hope that it will be useful,
 * but WITHOUT ANY WARRANTY; without even the implied warranty of
 * MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
 * European Union Public Licence for more details.
 *
 * You should have received a copy of the European Union Public Licence
 * along with this program. If not, see
 * http://ec.europa.eu/idabc/en/document/7330
 *
 * Änderungshistorie:
 * Datum      | Wer | Änderungsgrund
 * -------------------------------------------------------------------
 * 04.12.2007 | LUT | Erstellung
 * -------------------------------------------------------------------
 *
 * @author Christoph Lutz (D-III-ITD-D101)
 * @version 1.0
 * 
 */
package de.muenchen.allg.ooo;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.File;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.Collections;
import java.util.HashMap;
import java.util.List;

import com.sun.star.beans.PropertyValue;
import com.sun.star.beans.PropertyVetoException;
import com.sun.star.beans.UnknownPropertyException;
import com.sun.star.beans.XPropertyChangeListener;
import com.sun.star.beans.XPropertySet;
import com.sun.star.beans.XPropertySetInfo;
import com.sun.star.beans.XVetoableChangeListener;
import com.sun.star.container.ElementExistException;
import com.sun.star.container.NoSuchElementException;
import com.sun.star.document.XStorageBasedDocument;
import com.sun.star.embed.ElementModes;
import com.sun.star.embed.InvalidStorageException;
import com.sun.star.embed.StorageWrappedTargetException;
import com.sun.star.embed.XEncryptionProtectedSource;
import com.sun.star.embed.XStorage;
import com.sun.star.embed.XTransactedObject;
import com.sun.star.embed.XTransactionBroadcaster;
import com.sun.star.embed.XTransactionListener;
import com.sun.star.io.BufferSizeExceededException;
import com.sun.star.io.IOException;
import com.sun.star.io.NotConnectedException;
import com.sun.star.io.XInputStream;
import com.sun.star.io.XOutputStream;
import com.sun.star.io.XSeekable;
import com.sun.star.io.XStream;
import com.sun.star.lang.IllegalArgumentException;
import com.sun.star.lang.WrappedTargetException;
import com.sun.star.lang.XComponent;
import com.sun.star.lang.XEventListener;
import com.sun.star.packages.NoEncryptionException;
import com.sun.star.packages.WrongPasswordException;
import com.sun.star.uno.Type;

import de.muenchen.allg.afid.UNO;

/**
 * Diese Klasse definiert die notwendigen Services, um mit mit der Methode
 * XStorageBasedDocument.storeToStorage(...) XML-Daten aktiver Dokumente in den
 * Hauptspeicher zu speichern. Dazu wird mit {@link #createByteArrayStorage()}
 * ein Objekt erzeugt, das die Daten im Hauptspeicher verwaltet und aus
 * Java-Sicht besonders leicht weiter verarbeitet werden kann. Mit
 * {@link #createXStorage(de.muenchen.allg.ooo.MemoryStorage.ByteArrayStorage)}
 * wird ein entsprechender UNO-Service erzeugt, der für storeToStorage(...)
 * geeignet ist.
 * 
 * @author Christoph Lutz (D-III-ITD-D101)
 */
public class MemoryStorage {

	/**
	 * Diese Methode erzeugt einen ByteArrayStorage in dem Dokumente
	 * unkomprimiert und unverschlüsselt im Hauptspeicher abgelegt werden
	 * können.
	 * 
	 * @author Christoph Lutz (D-III-ITD-D101)
	 */
	public static ByteArrayStorage createByteArrayStorage() {
		return new ByteArrayStorage();
	}

	/**
	 * Erzeugt eine Implementierung des Service com::sun::star::embed::Storage,
	 * bei dem die Daten eines Dokuments in einem ByteArrayStorage abgelegt
	 * werden.
	 * 
	 * Achtung: Der UNO-Service definiert viele Methoden, die jedoch in der
	 * Praxis nicht alle relevant sind. Um toten Code zu vermeiden, und da die
	 * Spezifikation dieser Methoden in der UNO-API nicht aussagekräftig genug
	 * ist, wurden die unnötigen Methoden daher nicht implementiert. Das birgt
	 * aber auch das Risiko, dass neue OOo-Versionen in Zukunft das Storage
	 * anders ansprechen und dabei nicht implementierte Methoden verwenden
	 * könnten. Als Hilfestellung bei der Fehlersuche dient hier das
	 * public-field {@link #logger}, in dem ein Logger hinterlegt werden kann
	 * und über den verwendete, aber nicht implementierte Methoden erkannt
	 * werden können.
	 * 
	 * @author Christoph Lutz (D-III-ITD-D101)
	 */
	public static XStorage createXStorage(ByteArrayStorage bas) {
		return new StorageImpl(bas, "", ElementModes.WRITE);
	}

	/**
	 * Logger-Interfaces für die Behandlung von Debug-Ausgaben dieser Klasse.
	 * Für Debugging-Zwecke kann es sinnvoll und notwendig sein, einen eigenen
	 * Logger zu definieren und diesen im statischen Feld MemoryStorage.logger
	 * zu setzen.
	 */
	public interface Logger {
		public void log(String s);
	}

	/**
	 * Der Logger, auf den Meldungen ausgegeben werden, wenn er gesetzt ist.
	 * Derzeit nicht besonders sinnvoll, da private. Es muss eine Accessor-Methode
	 * eingeführt werden, um ihn zu setzen.
	 */
	private static Logger logger = null;

	/**
	 * Gibt eine Logger-Meldung mit der Nachricht s auf dem Logger aus, wenn
	 * dieser != null ist.
	 * 
	 * @author Christoph Lutz (D-III-ITD-D101)
	 */
	private static void log(String s) {
		if (logger != null)
			logger.log(s);
	}

	/**
	 * NotYetImplemented: Gibt eine Logger-Meldung mit der Nachricht
	 * "NotYetImplemented: " + s für eine Methode aus, die derzeit nicht
	 * implementiert ist.
	 * 
	 * @author Christoph Lutz (D-III-ITD-D101)
	 */
	private static void NYI(String s) {
		log("NotYetImplemented: " + s);
	}

	/**
	 * Diese Klasse definiert ein Storage, das benannte ByteArray-Datenblöcke
	 * aufnehmen kann. Das Storage enthält ausschließlich Elemente mit Daten und
	 * hält diese in einer flachen Struktur ohne Hierarchie. Daher kennt es
	 * keine Verzeichnisse, die einzelnen Elemente können aber Namen mit "/"
	 * enthalten.
	 * 
	 * @author Christoph Lutz (D-III-ITD-D101)
	 */
	public static class ByteArrayStorage {

		private HashMap<String, byte[]> mapNameToData;
		private HashMap<String, String> mapNameToMediaType;
		private String mediaType;

		private ByteArrayStorage() {
			this.mapNameToData = new HashMap<String, byte[]>();
			this.mapNameToMediaType = new HashMap<String, String>();
			mediaType = null;
		}

		/**
		 * Fügt das Element elementName mit den Daten data und dem MediaType
		 * mediaType in dieses Storage ein.
		 * 
		 * @author Christoph Lutz (D-III-ITD-D101)
		 */
		public void setData(String elementName, byte[] data, String mediaType) {
			mapNameToData.put(elementName, data);
			mapNameToMediaType.put(elementName, mediaType);
		}

		/**
		 * Liefert die Anzahl Datenbytes des Elements elementName zurück.
		 * 
		 * @author Christoph Lutz (D-III-ITD-D101)
		 */
		public int getSize(String elementName) {
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
		 * 
		 * @author Christoph Lutz (D-III-ITD-D101)
		 */
		public InputStream getInputStream(String elementName)
				throws java.util.NoSuchElementException {
			byte[] data = mapNameToData.get(elementName);
			if (data != null)
				return new ByteArrayInputStream(data);
			else
				throw new java.util.NoSuchElementException(elementName);
		}

		/**
		 * Liefert die namen aller Einträge des Storage in einer alphabetisch
		 * sortierten Liste zurück.
		 * 
		 * @author Christoph Lutz (D-III-ITD-D101)
		 */
		public List<String> getElementNames() {
			List<String> list = new ArrayList<String>(mapNameToData.keySet());
			Collections.sort(list);
			return list;
		}

		/**
		 * Diese Methode erlaubt das Setzen des MediaTypes dieses Storages. Ist
		 * der mediaType != null, so wird er beim Erzeugen des Manifests mittels
		 * {@link #createManifest()} auch im Manifest manifestiert.
		 * 
		 * @param mediaType
		 *            Der mediaType für Textdokumente ist z.B.
		 *            application/vnd.oasis.opendocument.text
		 * 
		 * @author Christoph Lutz (D-III-ITD-D101)
		 */
		public void setMediatype(String mediaType) {
			this.mediaType = mediaType;
		}

		/**
		 * Diese Methode erzeugt eine Datei META-INF/manifest.xml, die die
		 * Elemente des Storage beschreibt, und nimmt sie in das Storage mit
		 * auf.
		 * 
		 * @author Christoph Lutz (D-III-ITD-D101)
		 */
		public void createManifest() {
			String manifest = "";
			manifest += "<?xml version=\"1.0\" encoding=\"UTF-8\"?>\n";
			manifest += "<manifest:manifest xmlns:manifest=\"urn:oasis:names:tc:opendocument:xmlns:manifest:1.0\">\n";

			if (mediaType != null)
				manifest += "  <manifest:file-entry manifest:media-type=\""
						+ mediaType + "\" manifest:full-path=\"/\"/>\n";

			for (String name : mapNameToMediaType.keySet()) {
				String mediaType = mapNameToMediaType.get(name);
				manifest += "  <manifest:file-entry manifest:media-type=\""
						+ mediaType + "\" manifest:full-path=\"" + name
						+ "\"/>\n";
			}

			manifest += "</manifest:manifest>";
			mapNameToData.put("META-INF/manifest.xml", manifest.getBytes());
		}
	}

	/**
	 * Diese Klasse implementiert den UNO-Service
	 * com::sun::star::embed::Storage. Viele Methoden müssen definiert sein,
	 * damit der Service Storage von der UNO-Bridge korrekt anerkannt wird und
	 * ohne Fehler verwendet werden kann. In der Praxis sind aber viele der
	 * definierten Methoden überflüssig. Sie werden daher auch hier nicht
	 * implementiert, sonder nur definiert. Normalerweise sollte dies keine
	 * Probleme bereiten. Sollte es damit doch einmal Probleme geben, kann mit
	 * Hilfe des Loggers erkannt werden, ob eine nicht implementierte Methode
	 * verwendet wurde und diese implementiert werden.
	 * 
	 * @author Christoph Lutz (D-III-ITD-D101)
	 */
	private static class StorageImpl extends SimplePropertySet implements
			XStorage, XTransactedObject, XTransactionBroadcaster,
			XEncryptionProtectedSource {

		private final String namePrefix;
		private final ByteArrayStorage bas;
		private boolean isRoot;

		private StorageImpl(ByteArrayStorage bas, String namePrefix,
				int openMode) {
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

		public XStream cloneEncryptedStreamElement(String arg0, String arg1)
				throws InvalidStorageException, IllegalArgumentException,
				NoEncryptionException, WrongPasswordException, IOException,
				StorageWrappedTargetException {
			NYI(this + ".cloneEncryptedStreamElement" + arg0);
			return null;
		}

		public XStream cloneStreamElement(String arg0)
				throws InvalidStorageException, IllegalArgumentException,
				WrongPasswordException, IOException,
				StorageWrappedTargetException {
			NYI(this + ".cloneStreamElement " + arg0);
			return null;
		}

		public void copyElementTo(String arg0, XStorage arg1, String arg2)
				throws InvalidStorageException, IllegalArgumentException,
				NoSuchElementException, ElementExistException, IOException,
				StorageWrappedTargetException {
			NYI(this + ".copyElementTo " + arg0);
		}

		public void copyLastCommitTo(XStorage arg0)
				throws InvalidStorageException, IllegalArgumentException,
				IOException, StorageWrappedTargetException {
			NYI(this + ".copyLastCommitTo " + arg0);
		}

		public void copyStorageElementLastCommitTo(String arg0, XStorage arg1)
				throws InvalidStorageException, IllegalArgumentException,
				IOException, StorageWrappedTargetException {
			NYI(this + ".copyStorageElementLastCommitTo " + arg0);
		}

		public void copyToStorage(XStorage arg0)
				throws InvalidStorageException, IllegalArgumentException,
				IOException, StorageWrappedTargetException {
			NYI(this + ".copyToStorage " + arg0);
		}

		public boolean isStorageElement(String arg0)
				throws NoSuchElementException, IllegalArgumentException,
				InvalidStorageException {
			NYI(this + ".isStorageElement " + arg0);
			return false;
		}

		public boolean isStreamElement(String arg0)
				throws NoSuchElementException, IllegalArgumentException,
				InvalidStorageException {
			NYI(this + ".isStreamElement " + arg0);
			return false;
		}

		public void moveElementTo(String arg0, XStorage arg1, String arg2)
				throws InvalidStorageException, IllegalArgumentException,
				NoSuchElementException, ElementExistException, IOException,
				StorageWrappedTargetException {
			NYI(this + ".moveElementTo " + arg0);
		}

		public XStream openEncryptedStreamElement(String arg0, int arg1,
				String arg2) throws InvalidStorageException,
				IllegalArgumentException, NoEncryptionException,
				WrongPasswordException, IOException,
				StorageWrappedTargetException {
			NYI(this + ".openEncryptedStreamElement " + arg0);
			return null;
		}

		public XStorage openStorageElement(String arg0, int arg1)
				throws InvalidStorageException, IllegalArgumentException,
				IOException, StorageWrappedTargetException {

			log(this + ".openStorageElement " + arg0);
			StorageImpl s = new StorageImpl(bas, namePrefix + arg0 + "/", arg1);
			return s;
		}

		public XStream openStreamElement(String name, int openMode)
				throws InvalidStorageException, IllegalArgumentException,
				WrongPasswordException, IOException,
				StorageWrappedTargetException {

			log(this + ".openStreamElement " + name);
			return new StreamElementImpl(bas, namePrefix + name, openMode);
		}

		public void removeElement(String arg0) throws InvalidStorageException,
				IllegalArgumentException, NoSuchElementException, IOException,
				StorageWrappedTargetException {
			NYI(this + ".removeElement " + arg0);
		}

		public void renameElement(String arg0, String arg1)
				throws InvalidStorageException, IllegalArgumentException,
				NoSuchElementException, ElementExistException, IOException,
				StorageWrappedTargetException {
			NYI(this + ".renameElement " + arg0);
		}

		public Object getByName(String arg0) throws NoSuchElementException,
				WrappedTargetException {
			NYI(this + ".getByName " + arg0);
			return null;
		}

		public String[] getElementNames() {
			NYI(this + ".getElementNames");
			return new String[] {};
		}

		public boolean hasByName(String arg0) {
			log(this + ".hasByName " + arg0);
			for (String name : bas.getElementNames()) {
				name += "/";
				if (name.startsWith(namePrefix + arg0 + "/"))
					return true;
			}
			return false;
		}

		public Type getElementType() {
			NYI(this + ".getElementType");
			return null;
		}

		public boolean hasElements() {
			NYI(this + ".hasElements");
			return false;
		}

		public void addEventListener(XEventListener arg0) {
			NYI(this + ".addEventListener");
		}

		public void dispose() {
			NYI(this + ".dispose");
		}

		public void removeEventListener(XEventListener arg0) {
			NYI(this + ".removeEventListener");
		}

		public void commit() throws IOException, WrappedTargetException {
			log(this + ".commit");
			if (isRoot) {

				Object mediaType = props.get("MediaType");
				if (mediaType != null)
					bas.setMediatype(mediaType.toString());

				bas.createManifest();
			}
		}

		public void revert() throws IOException, WrappedTargetException {
			NYI(this + ".revert");
		}

		public void addTransactionListener(XTransactionListener arg0) {
			NYI(this + ".addTransactionListener");
		}

		public void removeTransactionListener(XTransactionListener arg0) {
			NYI(this + ".removeTransactionListener");
		}

		public void removeEncryption() throws IOException {
			NYI(this + ".removeEncryption");
		}

		public void setEncryptionPassword(String arg0) throws IOException {
			NYI(this + ".setEncryptionPassword");
		}

		public String toString() {
			return props.get("URL").toString();
		}
	}

	/**
	 * Diese Klasse implementiert den UNO-Service
	 * com::sun::star::embed::StorageStream. Viele Methoden müssen definiert
	 * sein, damit der Service StorageStream von der UNO-Bridge korrekt
	 * anerkannt wird und ohne Fehler verwendet werden kann. In der Praxis sind
	 * aber einige der definierten Methoden überflüssig. Sie werden daher auch
	 * nicht implementiert, sonder nur definiert. Normalerweise sollte dies
	 * keine Probleme bereiten. Sollte es damit doch einmal Probleme geben, kann
	 * mit Hilfe des Loggers erkannt werden, ob eine nicht implementierte
	 * Methode verwendet wurde und diese implementiert werden.
	 * 
	 * @author Christoph Lutz (D-III-ITD-D101)
	 */
	private static class StreamElementImpl extends SimplePropertySet implements
			XStream, XPropertySet, XSeekable, XEncryptionProtectedSource,
			XComponent {

		private final String name;

		private final ByteArrayStorage bas;

		private final String cn;

		private final ByteArrayOutputStream baos = new ByteArrayOutputStream();

		private final XOutputStream xOutputStream = new XOutputStream() {

			public void writeBytes(byte[] arg0) throws NotConnectedException,
					BufferSizeExceededException, IOException {

				log(cn + ".XOutputStream.writeBytes (length=" + arg0.length
						+ ")");
				try {
					baos.write(arg0);
				} catch (java.io.IOException e) {
					throw new IOException(e.toString());
				}
			}

			public void flush() throws NotConnectedException,
					BufferSizeExceededException, IOException {

				log(cn + ".XOutputStream.flush");
				flushIntern();
			}

			private void flushIntern() {
				bas.setData(name, baos.toByteArray(), getMediaType());
				props.put("Size", bas.getSize(name));
			}

			public void closeOutput() throws NotConnectedException,
					BufferSizeExceededException, IOException {

				log(cn + ".XOutputStream.closeOutput");
				flushIntern();
				try {
					baos.close();
				} catch (java.io.IOException e) {
					throw new IOException(e.toString());
				}
			}
		};

		private final XInputStream xInputStream = new XInputStream() {

			public void skipBytes(int arg0) throws NotConnectedException,
					BufferSizeExceededException, IOException {
				NYI(cn + ".XInputStream.skipBytes");
			}

			public int readSomeBytes(byte[][] arg0, int arg1)
					throws NotConnectedException, BufferSizeExceededException,
					IOException {
				NYI(cn + ".XInputStream.readSomeBytes");
				return 0;
			}

			public int readBytes(byte[][] arg0, int arg1)
					throws NotConnectedException, BufferSizeExceededException,
					IOException {
				NYI(cn + ".XInputStream.readBytes");
				return 0;
			}

			public void closeInput() throws NotConnectedException, IOException {
				NYI(cn + ".XInputStream.closeInput");
			}

			public int available() throws NotConnectedException, IOException {
				NYI(cn + ".XInputStream.available");
				return 0;
			}
		};

		public StreamElementImpl(ByteArrayStorage bas, String name, int openMode) {
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

		public XInputStream getInputStream() {
			log(this + ".getInputStream");
			return xInputStream;
		}

		private String getMediaType() {
			Object o = props.get("MediaType");
			return (o == null) ? "" : o.toString();
		}

		public XOutputStream getOutputStream() {
			log(this + ".getOutputStream");
			return xOutputStream;
		}

		public String toString() {
			return "streamelement:///" + name;
		}

		public long getLength() throws IOException {
			NYI(this + ".getLength");
			return 0;
		}

		public long getPosition() throws IOException {
			NYI(this + ".getPosition");
			return 0;
		}

		public void seek(long arg0) throws IllegalArgumentException,
				IOException {
			NYI(this + ".seek");
		}

		public void removeEncryption() throws IOException {
			NYI(this + ".removeEncryption");
		}

		public void setEncryptionPassword(String arg0) throws IOException {
			NYI(this + ".setEncryptionPassword");
		}

		public void addEventListener(XEventListener arg0) {
			NYI(this + ".addEventListener");
		}

		public void dispose() {
			NYI(this + ".dispose");
		}

		public void removeEventListener(XEventListener arg0) {
			NYI(this + ".removeEventListener");
		}

	}

	/**
	 * Diese Klasse implementiert ein einfaches PropertySet basierend auf einer
	 * HashMap. Auch hier sind viele in der Praxis nicht benötigte Methoden
	 * nicht implementiert, können aber bei Bedarf implementiert werden.
	 * 
	 * @author Christoph Lutz (D-III-ITD-D101)
	 */
	private static class SimplePropertySet implements XPropertySet {
		protected final HashMap<String, Object> props;

		public SimplePropertySet() {
			props = new HashMap<String, Object>();
		}

		public void addPropertyChangeListener(String arg0,
				XPropertyChangeListener arg1) throws UnknownPropertyException,
				WrappedTargetException {
			NYI(this + ".addPropertyChangeListener");
		}

		public void addVetoableChangeListener(String arg0,
				XVetoableChangeListener arg1) throws UnknownPropertyException,
				WrappedTargetException {
			NYI(this + ".addVetoableChangeListener");
		}

		public XPropertySetInfo getPropertySetInfo() {
			NYI(this + ".getPropertySetInfo");
			return null;
		}

		public Object getPropertyValue(String arg0)
				throws UnknownPropertyException, WrappedTargetException {

			log(this + ".getPropertyValue " + arg0);
			Object o = props.get(arg0);
			if (o != null)
				return o;
			else
				throw new UnknownPropertyException(arg0);
		}

		public void removePropertyChangeListener(String arg0,
				XPropertyChangeListener arg1) throws UnknownPropertyException,
				WrappedTargetException {
			NYI(this + ".removePropertyChangeListener " + arg0);
		}

		public void removeVetoableChangeListener(String arg0,
				XVetoableChangeListener arg1) throws UnknownPropertyException,
				WrappedTargetException {
			NYI(this + ".removeVetoableChangeListener " + arg0);
		}

		public void setPropertyValue(String arg0, Object arg1)
				throws UnknownPropertyException, PropertyVetoException,
				IllegalArgumentException, WrappedTargetException {

			log(this + ".setPropertyValue " + arg0 + " " + arg1);
			props.put(arg0, arg1);
		}
	}

	/**
	 * Testmethode: Speichert das aktive Vordergrunddokument in ein Storage und
	 * dieses nach /tmp/test.odt.
	 * 
	 * @author Christoph Lutz (D-III-ITD-D101)
	 */
	public static void main(String[] args) throws Exception {
		MemoryStorage.logger = new Logger() {
			public void log(String s) {
				System.out.println(s);
			}
		};

		UNO.init();
		XStorageBasedDocument sbd = UNO.XStorageBasedDocument(UNO.desktop
				.getCurrentComponent());

		log("UNO.init() done");

		long startTime = System.currentTimeMillis();

		final ByteArrayStorage bas = createByteArrayStorage();
		XStorage storage = createXStorage(bas);
		try {

			sbd.storeToStorage(storage, new PropertyValue[] {});

		} catch (Throwable t) {
			t.printStackTrace();
		}

		for (String name : bas.getElementNames()) {
			log("Element: " + name);
		}

		log("storeToStorage finished after "
				+ (System.currentTimeMillis() - startTime) + " ms");

		new ODFMerger.StorageZipOutput(new ODFMerger.Storage() {
			public InputStream getInputStream(String elementName)
					throws java.util.NoSuchElementException {
				return bas.getInputStream(elementName);
			}

			public List<String> getElementNames() {
				return bas.getElementNames();
			}
		}).writeToFile(new File("/tmp/test.odt"));

		System.exit(0);
	}
}
