/*
 * Dateiname: ODFMerger.java
 * Projekt  : UNOHelper
 * Funktion : Merged ODF-Dokumente in ein Ergebnisdokument zusammen.
 * 
 * Copyright (c) 2008 Landeshauptstadt München
 *
 * This program is free software: you can redistribute it and/or modify
 * it under the terms of the European Union Public Licence (EUPL),
 * version 1.0 (or any later version).
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
 * 03.12.2007 | LUT | Erstellung
 * -------------------------------------------------------------------
 *
 * @author Christoph Lutz (D-III-ITD-D101)
 * @version 1.0
 * 
 */
package de.muenchen.allg.ooo;

import java.io.ByteArrayInputStream;
import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.StringWriter;
import java.util.ArrayList;
import java.util.Enumeration;
import java.util.HashMap;
import java.util.HashSet;
import java.util.List;
import java.util.Map;
import java.util.NoSuchElementException;
import java.util.zip.ZipEntry;
import java.util.zip.ZipException;
import java.util.zip.ZipFile;
import java.util.zip.ZipOutputStream;

import javax.xml.parsers.DocumentBuilder;
import javax.xml.parsers.DocumentBuilderFactory;
import javax.xml.transform.Transformer;
import javax.xml.transform.TransformerFactory;
import javax.xml.transform.dom.DOMSource;
import javax.xml.transform.stream.StreamResult;

import org.w3c.dom.Document;
import org.w3c.dom.NamedNodeMap;
import org.w3c.dom.Node;
import org.w3c.dom.NodeList;

/**
 * Definiert einen ODFMerger, der auf XML-Basis mehrere Dokumente zu einem
 * Gesamtdokument zusammen fassen kann. Der Merger setzt aktuell vorraus, dass die zu
 * mergenden Dokumente gleichartig sind, bzw. aus dem selben Ursprungsdokument
 * erzeugt wurden. Mit dieser Voraussetzung ist es nicht notwendig, bestimmte
 * Ressourcen (z.B. Bilder und manche Styles) aufzudoppeln - sie werden einfach
 * mehrfach verwendet. Der Merger ist also nicht für beliebige unterschiedliche
 * Dokumente geeignet, sollte aber z.B. für die Erzeugung von Gesamtdokumenten in
 * Rahmen eines WollMux-Komfortdrucks alle notwendigen Elemente übernehmen und ggf.
 * mit Fixup-Methoden anpassen.
 * 
 * @author Christoph Lutz (D-III-ITD-D101)
 */
public class ODFMerger
{

  /**
   * Enthält den Merger mit der entsprechenden Merge-Logik für meta.xml
   */
  MetaMerger mm;

  /**
   * Enthält den Merger mit der entsprechenden Merge-Logik für styles.xml
   */
  StylesMerger sm;

  /**
   * Enthält den Merger mit der entsprechenden Merge-Logik für content.xml
   */
  ContentMerger cm;

  /**
   * Das zuerst hinzugefügte Dokument hat eine Sonderrolle, da es als Basis für das
   * Erbebnisdokument verwendet wird. Angepasst werden lediglich styles.xml,
   * content.xml und meta.xml. Alle anderen Ressourcen werden aus dem ersten
   * hinzugefügten Storage übernommen.
   */
  Storage firstStorage;

  /**
   * Definiert ein Storage für ein ODF-Dokument.
   * 
   * @author Christoph Lutz (D-III-ITD-D101)
   */
  public interface Storage
  {
    InputStream getInputStream(String elementName) throws NoSuchElementException;

    List<String> getElementNames();
  }

  /**
   * Erzeugt einen neuen ODFMerger, in den später per add Dokumente in Form von
   * Storages aufgenommen werden können.
   */
  public ODFMerger()
  {
    mm = new MetaMerger();
    cm = new ContentMerger();
    sm = new StylesMerger();
  }

  /**
   * Fügt das Dokument in Form des Storages s zum Ergebnisdokument hinzu.
   * 
   * @author Christoph Lutz (D-III-ITD-D101)
   */
  public void add(Storage s)
  {
    if (firstStorage == null) firstStorage = s;

    try
    {
      int pageOffset = mm.getPageCountConsiderFillerPages();
      mm.add(s.getInputStream("meta.xml"));
      sm.add(s.getInputStream("styles.xml"), pageOffset);
      cm.add(s.getInputStream("content.xml"), pageOffset,
        sm.getLastMasterStylesChangeMap());
    }
    catch (NoSuchElementException e)
    {
      e.printStackTrace();
    }
  }

  /**
   * Liefert ein Storage zurück, das das Ergebnisdokument des bisherigen Merges
   * repräsentiert.
   * 
   * @author Christoph Lutz (D-III-ITD-D101)
   */
  public Storage getResultStorage()
  {
    if (firstStorage != null)
    {
      OverrideStorage os = new OverrideStorage(firstStorage);
      os.override("meta.xml", mm.getResultData());
      os.override("styles.xml", sm.getResultData());
      os.override("content.xml", cm.getResultData());
      return os;
    }
    return null;
  }

  /**
   * Merged die Daten der meta.xml-Dateien, die in den mit add hinzugefügten Storages
   * vorhanden sind.
   * 
   * @author Christoph Lutz (D-III-ITD-D101)
   */
  public static class MetaMerger extends Merger
  {

    private int pageCount = 0;

    private HashMap<String, Integer> counters = new HashMap<String, Integer>();

    /**
     * Merged die in is mitgelieferte meta.xml in das Ergebnisdokument ein.
     * 
     * @author Christoph Lutz (D-III-ITD-D101) TODO: testen
     */
    public void add(InputStream is)
    {
      DocumentBuilderFactory factory = DocumentBuilderFactory.newInstance();
      DocumentBuilder docBuilder;

      Document doc = null;
      try
      {
        docBuilder = factory.newDocumentBuilder();
        doc = docBuilder.parse(is);
      }
      catch (Exception e)
      {
        e.printStackTrace();
      }

      Node statistic =
        getFirstChild(getFirstChild(getFirstChild(doc), "office:meta"),
          "meta:document-statistic");
      if (statistic != null)
      {
        NamedNodeMap atts = statistic.getAttributes();
        for (int i = 0; i < atts.getLength(); ++i)
        {
          Node att = atts.item(i);
          String name = att.getNodeName();
          String value = att.getFirstChild().getNodeValue();
          if (name.endsWith("count"))
          {
            Integer oldValue = counters.get(name);
            if (oldValue == null) oldValue = Integer.valueOf(0);
            try
            {
              Integer newValue = Integer.valueOf(value);
              counters.put(name, oldValue + newValue);
              if ("meta:page-count".equals(name))
                pageCount = getPageCountConsiderFillerPages() + newValue;
            }
            catch (NumberFormatException e)
            {
              e.printStackTrace();
            }
          }
        }
      }
    }

    /*
     * (non-Javadoc)
     * 
     * @see de.muenchen.allg.itd51.test.DOMBasedODFMerger.Merger#getResultData()
     */
    public byte[] getResultData()
    {
      // TODO Die in counters und in pageCount gesammelten Werte auch
      // tatsächlich in das Ergebnisdokument schreiben.
      return super.getResultData();
    }

    /**
     * Liefert die aktuelle Seitenzahl des gemergten Dokuments mit bereits
     * eingefügten bzw. noch einzufügenden Leerseiten damit das nächste Dokument auf
     * einer ungeraden Seite starten kann.
     * 
     * @author Christoph Lutz (D-III-ITD-D101)
     */
    public int getPageCountConsiderFillerPages()
    {
      return pageCount + (pageCount % 2);
    }
  }

  /**
   * Merged die Daten der styles.xml-Dateien, die in den mit add hinzugefügten
   * Storages vorhanden sind.
   * 
   * @author Christoph Lutz (D-III-ITD-D101)
   */
  public static class StylesMerger extends Merger
  {
    private HashMap<String, String> masterStylesChangeMap =
      new HashMap<String, String>();

    public StylesMerger()
    {
      toReplaceByChildNodeNames.add("text:page-count");
    }

    /**
     * Merged die in is mitgelieferte styles.xml in das Ergebnisdokument ein.
     * 
     * @author Christoph Lutz (D-III-ITD-D101) TODO: testen
     */
    public void add(InputStream is, int pageOffset)
    {
      DocumentBuilderFactory factory = DocumentBuilderFactory.newInstance();
      DocumentBuilder docBuilder;

      try
      {
        docBuilder = factory.newDocumentBuilder();
        Document doc = docBuilder.parse(is);

        // remove Page-Count Felder
        deleteOrReplaceByChild(doc, toDeleteNodeNames, toReplaceByChildNodeNames);

        // unify masterStyleNames
        Node masterStyles =
          getFirstChild(doc.getFirstChild(), "office:master-styles");
        masterStylesChangeMap.clear();
        for (String s : getStyleNames(masterStyles))
          masterStylesChangeMap.put(s, s + "_" + pageOffset);
        adjustStyleNames(doc, masterStylesChangeMap, new String[] {
          "style:name", "style:master-page-name" });

        if (resultDoc == null)
        {
          resultDoc = (Document) doc.cloneNode(true);
        }
        else
        {
          // master Styles mergen
          Node srcNode = getFirstChild(doc.getFirstChild(), "office:master-styles");
          Node resultNode =
            getFirstChild(resultDoc.getFirstChild(), "office:master-styles");
          NodeList list = srcNode.getChildNodes();
          for (int i = 0; i < list.getLength(); ++i)
          {
            Node n = list.item(i);
            resultNode.appendChild(resultDoc.importNode(n, true));
          }
        }
      }
      catch (Exception e)
      {
        e.printStackTrace();
      }
    }

    public Map<String, String> getLastMasterStylesChangeMap()
    {
      return masterStylesChangeMap;
    }
  }

  /**
   * Merged die Daten der content.xml-Dateien, die in den mit add hinzugefügten
   * Storages vorhanden sind.
   * 
   * @author Christoph Lutz (D-III-ITD-D101)
   */
  public static class ContentMerger extends Merger
  {

    public ContentMerger()
    {
      toReplaceByChildNodeNames.add("text:drop-down");
      toReplaceByChildNodeNames.add("text:text-input");
      toReplaceByChildNodeNames.add("text:database-display");
      toReplaceByChildNodeNames.add("text:date");

      toDeleteNodeNames.add("office:annotation");
      toDeleteNodeNames.add("text:bookmark-start");
      toDeleteNodeNames.add("text:bookmark-end");
    }

    /**
     * Merged die in is mitgelieferte content.xml in das Ergebnisdokument ein.
     * 
     * @author Christoph Lutz (D-III-ITD-D101) TODO: testen
     */
    public void add(InputStream is, int pageOffset,
        Map<String, String> masterStylesChangeMap)
    {
      DocumentBuilderFactory factory = DocumentBuilderFactory.newInstance();
      DocumentBuilder docBuilder;

      try
      {
        docBuilder = factory.newDocumentBuilder();
        Document doc = docBuilder.parse(is);

        // remove TextFields und andere Elemente
        deleteOrReplaceByChild(doc, toDeleteNodeNames, toReplaceByChildNodeNames);

        // add page break to first text:p paragraph
        Node autoStyles =
          getFirstChild(doc.getFirstChild(), "office:automatic-styles");
        HashSet<String> names = getStyleNames(autoStyles);
        Node firstTextPNode = getFirstTextPNode(doc);
        String newParStyleName = createUnusedParStyleName(names);
        String oldParStyleName = adjustParStyleName(firstTextPNode, newParStyleName);
        String masterPageName = getMasterPageName(autoStyles, oldParStyleName);
        autoStyles.appendChild(createPageBreakParStyleNode(doc, newParStyleName,
          oldParStyleName, masterPageName));

        // addPageAnchorOffset
        addPageAnchorOffset(doc, pageOffset);

        // unify autoStyleNames
        HashMap<String, String> stylesChangeMap = new HashMap<String, String>();
        for (String s : getStyleNames(autoStyles))
          stylesChangeMap.put(s, s + "_" + pageOffset);
        adjustStyleNames(doc, stylesChangeMap, new String[] {
          "style:name", "draw:style-name", "text:style-name",
          "presentation:style-name", "style:style-name", "style:parent-style-name",
          "chart:style-name", "table:style-name" });

        // unify masterStyleNames
        adjustStyleNames(doc, masterStylesChangeMap,
          new String[] { "style:master-page-name" });

        if (resultDoc == null)
        {
          resultDoc = (Document) doc.cloneNode(true);
        }
        else
        {
          // Automatic Styles mergen
          Node srcNode =
            getFirstChild(doc.getFirstChild(), "office:automatic-styles");
          Node resultNode =
            getFirstChild(resultDoc.getFirstChild(), "office:automatic-styles");
          NodeList list = srcNode.getChildNodes();
          for (int i = 0; i < list.getLength(); ++i)
          {
            Node n = list.item(i);
            resultNode.appendChild(resultDoc.importNode(n, true));
          }

          // office:text mergen
          // - sequence-decls werden nicht gemerged.
          // - An der Seite verankerte Objekte stehen dabei in einem
          // Block am Anfang
          // - die restlichen Elemente werden jeweils hinten
          // angehängt.
          srcNode =
            getFirstChild(getFirstChild(doc.getFirstChild(), "office:body"),
              "office:text");
          resultNode =
            getFirstChild(getFirstChild(resultDoc.getFirstChild(), "office:body"),
              "office:text");
          Node lastPageAnchoredNode = getLastPageAnchoredNode(resultNode);
          Node beginContentBlock = null;
          if (lastPageAnchoredNode != null)
            beginContentBlock = lastPageAnchoredNode.getNextSibling();
          list = srcNode.getChildNodes();
          for (int i = 0; i < list.getLength(); ++i)
          {
            Node n = list.item(i);
            if ("text:sequence-decls".equals(n.getNodeName())) continue;
            if (isPageAnchored(n))
            {
              if (beginContentBlock != null)
              {
                resultNode.insertBefore(resultDoc.importNode(n, true),
                  beginContentBlock);
              }
            }
            else
              resultNode.appendChild(resultDoc.importNode(n, true));
          }

        }
      }
      catch (Exception e)
      {
        e.printStackTrace();
      }
    }

    /**
     * Diese Methode liefert den Wert des Attributs style:master-page-name zurück,
     * das im Style styleName oder in Styles, von denen styleName abhängig ist,
     * definiert ist; Gesucht wird im Abschnitt AutoStyles, dessen Knoten autoStyles
     * als Übergabeparameter erwartet wird; Ist keine MasterPage für einen dieser
     * Styles definiert oder ist der gesuchte Style selbst nicht definiert, so wird
     * "Standard" zurück geliefert.
     * 
     * @param autoStyles
     * @param styleName
     * 
     * @author Christoph Lutz (D-III-ITD-D101)
     */
    private String getMasterPageName(Node autoStyles, String styleName)
    {
      NodeList list = autoStyles.getChildNodes();
      for (int i = 0; i < list.getLength(); ++i)
      {
        Node n = list.item(i);
        NamedNodeMap nnm = n.getAttributes();
        if (nnm != null
          && styleName.equals(nnm.getNamedItem("style:name").getFirstChild().getNodeValue()))
        {

          Node mpn = nnm.getNamedItem("style:master-page-name");
          if (mpn != null) return mpn.getFirstChild().getNodeValue();

          Node psn = nnm.getNamedItem("style:parent-style-name");
          if (psn != null)
            return getMasterPageName(autoStyles, psn.getFirstChild().getNodeValue());

          return "Standard";
        }
      }
      return "Standard";
    }

    /**
     * Erzeugt die Style-Definition eines Styles newParStyleName, der von
     * oldParStyleName erbt und einen Seitennumbruch an diesem Paragraphen
     * verursacht, dessen Seitenzähler mit 1 initialisiert ist. Das Resultat in
     * XML-Syntax: <style:style style:name="$newParStyleName"
     * style:family="paragraph" style:parent-style-name="$oldParStyleName"
     * style:master-page-name="$oldMasterStyleName"><style:paragraph-properties
     * style:page-number="1"/></style:style>
     * 
     * @param doc
     *          Das Dokument in das die Style-Definition später eingepflegt werden
     *          soll.
     * @param newParStyleName
     *          der Name der zu erzeugenden Style-Definition.
     * @param oldParStyleName
     *          der Name des Styles, von dem geerbt wird.
     * @param oldMasterPageStyleName
     *          der Style-Name der MasterPage, die als Seitenvorlage für die neue
     *          Seite verwendet werden soll. Soll keine spezielle Seite gesetzt
     *          werden, so sorgt die Übergabe von "Standard" für das gewünschte
     *          Ergebnis. Der Wert ist verpflichtend, da ohne ihn kein Seitenumbruch
     *          definiert ist.
     * 
     * @author Christoph Lutz (D-III-ITD-D101)
     */
    private Node createPageBreakParStyleNode(Document doc, String newParStyleName,
        String oldParStyleName, String oldMasterPageStyleName)
    {
      Node styleStyle = doc.createElement("style:style");
      Node att = doc.createAttribute("style:name");
      att.appendChild(doc.createTextNode(newParStyleName));
      styleStyle.getAttributes().setNamedItem(att);
      att = doc.createAttribute("style:family");
      att.appendChild(doc.createTextNode("paragraph"));
      styleStyle.getAttributes().setNamedItem(att);
      att = doc.createAttribute("style:parent-style-name");
      att.appendChild(doc.createTextNode(oldParStyleName));
      styleStyle.getAttributes().setNamedItem(att);
      att = doc.createAttribute("style:master-page-name");
      att.appendChild(doc.createTextNode(oldMasterPageStyleName));
      styleStyle.getAttributes().setNamedItem(att);

      Node parProp = doc.createElement("style:paragraph-properties");
      att = doc.createAttribute("style:page-number");
      att.appendChild(doc.createTextNode("1"));
      parProp.getAttributes().setNamedItem(att);
      styleStyle.appendChild(parProp);
      return styleStyle;
    }

    /**
     * Wenn das übergebene Element node ein Attribut text:style-name definiert, so
     * wird der Wert dieses Attributs neu mit newParStyleName belegt. Diese Methode
     * dient üblicherweise dazu, den ersten Absatz eines Dokuments mit einer
     * Absatzformatierung zu belegen, in der ein Seitenumbruch definiert ist. Auf
     * diese Art wird der gewünschte Seitenumbruch vor jedes zu mergende Dokument
     * eingefügt.
     * 
     * @author Christoph Lutz (D-III-ITD-D101)
     */
    public static String adjustParStyleName(Node node, String newParStyleName)
    {
      NamedNodeMap atts = node.getAttributes();
      if (atts != null)
      {
        Node item = atts.getNamedItem("text:style-name");
        if (item != null)
        {
          String oldVal = item.getFirstChild().getNodeValue();
          item.getFirstChild().setNodeValue(newParStyleName);
          return oldVal;
        }
      }
      return null;
    }

    /**
     * Prüft, ob das übergebene Element (üblicherweise ein Text-Frame) an der Seite
     * verankert ist. Hintergrund ist der, dass bei der Suche nach dem ersten
     * Paragraph des Dokuments alle Text-Frames ignoriert werden müssen, die nicht
     * Teil des Haupttextteils des Dokuments sind. Der erste Paragraph des
     * Haupttextteils wird benötigt, um den Seitenumbruch vor dem hinzumergen eines
     * neuen Dokuments einzufügen.
     * 
     * @author Christoph Lutz (D-III-ITD-D101)
     */
    private static boolean isPageAnchored(Node node)
    {
      NamedNodeMap atts = node.getAttributes();
      if (atts != null)
      {
        Node anchorType = atts.getNamedItem("text:anchor-type");
        if (anchorType != null
          && "page".equals(anchorType.getFirstChild().getNodeValue())) return true;
      }
      return false;
    }

    /**
     * Im XML-Dokument müssen alle an der Seite verankerten Objekte in einem Block
     * hintereinander aufgeführt sein, damit OOo sie korrekt darstellen kann (ich
     * habe nicht geprüft, ob diese Anforderung aus der ODF-Spezifikation resultiert,
     * oder ob es sich hier um einen OOo-Bug handelt). Diese Methode liefert den
     * Knoten des letzten Objekts, das an der Seite verankert ist, zurück oder null,
     * wenn kein an der Seite verankertes Objekt im Dokument gefunden wurde.
     * 
     * @param officeTextNode
     *          Der Knoten des Elements office:text unterhalb dessen die an der Seite
     *          verankerten Objekte aufgelistet sind.
     * @return den Knoten des letzten Elements, das an der Seite verankert ist oder
     *         null, wenn kein solches Objekt vorhanden ist.
     * 
     * @author Christoph Lutz (D-III-ITD-D101)
     */
    private static Node getLastPageAnchoredNode(Node officeTextNode)
    {
      NodeList list = officeTextNode.getChildNodes();
      Node last = null;
      for (int i = 0; i < list.getLength(); ++i)
      {
        Node n = list.item(i);
        if (isPageAnchored(n)) last = n;
      }
      return last;
    }

    /**
     * Sucht im Dokument node rekursiv nach dem ersten Text-Paragraphen. Bei diesem
     * Paragraphen kann später der einzufügende Seitenumbruch angewendet werden.
     * Dabei ist wichtig, dass der Paragraph im Haupttextteil liegt. So müssen z.B.
     * Paragraphen in TextFrames, die an der Seite verankert sind, ignoriert werden.
     * Dies wird erreicht, in dem alle an der Seite verankerten Objekte gleich zu
     * Beginn ignoriert werden.
     * 
     * @param node
     *          Hier kann der Hauptknoten des Dokument angegeben werden - es wird
     *          rekursiv gesucht.
     * @return den ersten Paragraphen des Haupttextteils oder null, wenn keiner
     *         gefunden wurde.
     * 
     * @author Christoph Lutz (D-III-ITD-D101)
     */
    private static Node getFirstTextPNode(Node node)
    {
      if (isPageAnchored(node)) return null;

      if ("text:p".equals(node.getNodeName())) return node;

      // process childs
      NodeList list = node.getChildNodes();
      for (int i = 0; i < list.getLength(); i++)
      {
        Node firstTextPNode = getFirstTextPNode(list.item(i));
        if (firstTextPNode != null) return firstTextPNode;
      }
      return null;
    }

    /**
     * Diese Fixup-Methode korrigiert rekursiv alle an der Seite verankeren Objekte
     * in dem der neue Seiten-Offset innerhalb des gemergten Dokuments offset zu der
     * bereits gesetzten Startseite hinzugefügt wird.
     * 
     * @param node
     *          Den Knoten eines Objekts. Ist dieses Objekt an der Seite verankert,
     *          so wird der Offset offset zu dem bereits gesetzten Wert des Attributs
     *          text:anchor-page-number addiert.
     * @param offset
     *          Die Seitennummer der Seite, mit der das aktuelle Dokument im
     *          gemergten Dokument beginnt.
     * 
     * @author Christoph Lutz (D-III-ITD-D101)
     */
    private static void addPageAnchorOffset(Node node, int offset)
    {
      NamedNodeMap atts = node.getAttributes();
      if (atts != null)
      {
        Node item = atts.getNamedItem("text:anchor-type");
        if (item != null && "page".equals(item.getFirstChild().getNodeValue()))
        {
          item = atts.getNamedItem("text:anchor-page-number");
          String nStr = "1";
          if (item != null)
          {
            nStr = item.getFirstChild().getNodeValue();
            int anchorPageNumber = 1;
            try
            {
              anchorPageNumber = Integer.valueOf(nStr);
            }
            catch (NumberFormatException e)
            {
              e.printStackTrace();
            }
            item.getFirstChild().setNodeValue("" + (anchorPageNumber + offset));
          }
        }
      }

      // process childs
      NodeList list = node.getChildNodes();
      for (int i = 0; i < list.getLength(); i++)
      {
        addPageAnchorOffset(list.item(i), offset);
      }
    }
  }

  /**
   * Enthält gemeinsame Eigenschaften aller Merger.
   * 
   * @author Christoph Lutz (D-III-ITD-D101)
   */
  public static class Merger
  {

    /**
     * Das Ergebnisdokument, das mit dem ersten hinzugefügten Dokument initialisiert
     * wird und in dem später alle weiteren Dokumente zusammengemergt werden.
     */
    protected Document resultDoc = null;

    /**
     * Enthält die Namen aller XML-Elemente, die im Ergebnisdokument nicht erwünscht
     * sind und durch ihre Kinder ersetzt werden sollen (falls Kinder vorhanden
     * sind). Beispiel: Textfelder werden durch ihre Repräsentation ersetzt.
     */
    protected HashSet<String> toReplaceByChildNodeNames = null;

    /**
     * Enthält die Namen aller Knoten, die im Ergebnisdokument nicht erwünscht sind
     * und vollständig entfernt werden können. Z.B. Notizen oder Bookmarks. Damit
     * wird das Ergebnisdokument etwas entschlankt und nicht benötigte Elemente
     * müssen nicht unnötig angepasst (Fixups) werden.
     */
    protected HashSet<String> toDeleteNodeNames = null;

    public Merger()
    {
      this.toReplaceByChildNodeNames = new HashSet<String>();
      this.toDeleteNodeNames = new HashSet<String>();
    }

    /**
     * Sucht im Knoten parent nach Style-Definitionen und liefert die Namen aller
     * gefundenen Style-Definitionen als Menge zurück. Als parent wird dabei
     * üblicherweise ein Block text:auto-styles (aus content.xml) oder
     * text:master-styles (aus styles.xml) erwartet.
     * 
     * @author Christoph Lutz (D-III-ITD-D101)
     */
    protected static HashSet<String> getStyleNames(Node parent)
    {
      HashSet<String> names = new HashSet<String>();
      NodeList list = parent.getChildNodes();
      for (int i = 0; i < list.getLength(); ++i)
      {
        Node n = list.item(i);
        NamedNodeMap nnm = n.getAttributes();
        if (nnm != null)
        {
          names.add(nnm.getNamedItem("style:name").getFirstChild().getNodeValue());
        }
      }
      return names;
    }

    /**
     * Erzeugt den Namen eines neuen, garantiert unbenutzen Paragraph-Styles, der
     * bisher nicht in names vorhanden ist. Der Name von Paragraph-Styles ist
     * übelicherweise wie "P<zahl>" aufgebaut.
     * 
     * @param names
     *          Die Menge der bereits bekannten Paragraph-Styles
     * @return einen neuen unbenutzen Stylenamen nach dem Schame "P<zahl>".
     * 
     * @author Christoph Lutz (D-III-ITD-D101)
     */
    protected static String createUnusedParStyleName(HashSet<String> names)
    {
      for (int i = 1; true; ++i)
      {
        String n = "P" + i;
        if (!names.contains(n)) return n;
      }
    }

    /**
     * Diese Methode entfernt rekursiv Knoten aus dem Dokument, wobei toDelete alle
     * Elemente beschreibt, die komplett (mitsamt Kinder) gelöscht werden sollen und
     * toReplace alle Knoten beschreibt, die durch ihre Kinder ersetzt werden sollen.
     * 
     * @param node
     *          Der Knoten, ab dem rekursiv ersetzt werden soll.
     * @param toDelete
     *          beschreibt die Namen von Elementen, die komplett gelöscht werden
     *          sollen.
     * @param toReplace
     *          beschreibt die Namen von Elemente, die durch ihre Kinder ersetzt
     *          werden sollen.
     * 
     * @author Christoph Lutz (D-III-ITD-D101)
     */
    protected static void deleteOrReplaceByChild(Node node,
        HashSet<String> toDelete, HashSet<String> toReplace)
    {
      Node child = node.getFirstChild();
      while (child != null)
      {
        Node nextChild = child.getNextSibling();
        deleteOrReplaceByChild(child, toDelete, toReplace);
        child = nextChild;
      }

      String name = node.getNodeName();

      if (toDelete.contains(name))
      {
        node.getParentNode().removeChild(node);
      }

      else if (toReplace.contains(name))
      {
        NodeList children = node.getChildNodes();
        for (int i = 0; i < children.getLength(); ++i)
          node.getParentNode().insertBefore(children.item(i), node);
        node.getParentNode().removeChild(node);
      }
    }

    /**
     * Fixup-Methode für Stylenamen, die rekursiv Namen von Styles in Attributen
     * anpasst; Dabei beschreibt stylesChangeMap die Zuordnung von alten Namen auf
     * neue Namen und attributesToAdjust die Liste der Attribute, die angepasst
     * werden sollen, wenn sie in einem Element gesetzt sine.
     * 
     * @param node
     *          Ausgangsknoten für die rekursive Anpassung.
     * @param stylesChangeMap
     *          HashMap mit einer Zuordnung von alten Style-Namen auf die neu zu
     *          setzenden Style-Namen.
     * @param attributesToAdjust
     *          Eine Liste der Attribute, die in allen Elementen und Unterelementen
     *          angepasst werden sollen, wenn sie gesetzt sind.
     * 
     * @author Christoph Lutz (D-III-ITD-D101)
     */
    protected static void adjustStyleNames(Node node,
        Map<String, String> stylesChangeMap, String[] attributesToAdjust)
    {
      NamedNodeMap atts = node.getAttributes();
      if (atts != null)
      {
        for (String attName : attributesToAdjust)
        {
          Node item = atts.getNamedItem(attName);
          if (item != null)
          {
            String newVal = stylesChangeMap.get(item.getFirstChild().getNodeValue());
            if (newVal != null) item.getFirstChild().setNodeValue(newVal);
          }
        }
      }

      // process childs
      NodeList list = node.getChildNodes();
      for (int i = 0; i < list.getLength(); i++)
      {
        adjustStyleNames(list.item(i), stylesChangeMap, attributesToAdjust);
      }
    }

    /**
     * Liefert das XML-Ergebnisdokument des Mergevorgangs als byte[].
     * 
     * @author Christoph Lutz (D-III-ITD-D101)
     */
    public byte[] getResultData()
    {
      if (resultDoc != null) try
      {
        StringWriter sw = new StringWriter();
        TransformerFactory tf = TransformerFactory.newInstance();
        Transformer transformer = tf.newTransformer();
        transformer.transform(new DOMSource(resultDoc), new StreamResult(sw));
        return sw.toString().getBytes();
      }
      catch (Exception e)
      {
        e.printStackTrace();
      }
      return null;
    }

    /**
     * Liefert das erste Kind von parent mit dem Elementnamen element zurück
     * 
     * @author Christoph Lutz (D-III-ITD-D101)
     */
    protected static Node getFirstChild(Node parent, String element)
    {
      if (parent == null) return null;
      NodeList nl = parent.getChildNodes();
      for (int i = 0; i < nl.getLength(); i++)
      {
        Node child = nl.item(i);
        if (element.equals(child.getNodeName())) return child;
      }
      return null;
    }

    /**
     * Liefert parent.getFirstChild() zurück oder null, wenn parent selbst null ist.
     * 
     * @author Christoph Lutz (D-III-ITD-D101)
     */
    protected static Node getFirstChild(Node parent)
    {
      if (parent == null) return null;
      return parent.getFirstChild();
    }

    /**
     * Erzeugt eine simple Baumansicht des Knotens node und aller Unterknoten auf
     * System.out für Debugzwecke.
     * 
     * @author Christoph Lutz (D-III-ITD-D101)
     */
    protected static void dumpNode(Node node, String fillers)
    {
      System.out.println(fillers + node.getNodeName());
      NodeList list = node.getChildNodes();
      for (int i = 0; i < list.getLength(); ++i)
        dumpNode(list.item(i), fillers + "  ");
    }
  }

  /**
   * Repräsentiert ein Storage für ein normales gezipptes ODF-Dokument.
   * 
   * @author Christoph Lutz (D-III-ITD-D101)
   */
  private static class ZipedODFFileStorage implements Storage
  {

    private ZipFile zipFile;

    public ZipedODFFileStorage(File file) throws ZipException, IOException
    {
      this.zipFile = new ZipFile(file);
    }

    public InputStream getInputStream(String elementName)
    {
      try
      {
        return zipFile.getInputStream(zipFile.getEntry(elementName));
      }
      catch (IOException e)
      {
        throw new NoSuchElementException(elementName);
      }
    }

    public List<String> getElementNames()
    {
      List<String> names = new ArrayList<String>();
      Enumeration<? extends ZipEntry> entries = zipFile.entries();
      while (entries.hasMoreElements())
      {
        ZipEntry entry = entries.nextElement();
        names.add(entry.getName());
      }
      return names;
    }
  }

  /**
   * Repräsentiert ein Storage, das Ressourcen von einem anderen Storage übernimmt,
   * in dem jedoch einzelne Ressourcen verändert bzw. überschrieben werden können.
   * 
   * @author Christoph Lutz (D-III-ITD-D101)
   */
  private static class OverrideStorage implements Storage
  {
    private Storage storage;

    private HashMap<String, byte[]> overrideMap;

    public OverrideStorage(Storage storage)
    {
      this.storage = storage;
      this.overrideMap = new HashMap<String, byte[]>();
    }

    public void override(String elementName, byte[] data)
    {
      overrideMap.put(elementName, data);
    }

    public InputStream getInputStream(String elementName)
        throws NoSuchElementException
    {
      byte[] data = overrideMap.get(elementName);
      if (data != null)
      {
        return new ByteArrayInputStream(data);
      }
      return storage.getInputStream(elementName);
    }

    public List<String> getElementNames()
    {
      return storage.getElementNames();
    }
  }

  /**
   * Erweitert ein Storage um die Fähigkeit, die enthaltenen Daten in ein gezipptes
   * odf-File zu schreiben.
   * 
   * @author Christoph Lutz (D-III-ITD-D101)
   */
  public static class StorageZipOutput
  {
    private Storage storage;

    public StorageZipOutput(Storage storage)
    {
      this.storage = storage;
    }

    public void writeToFile(File file) throws IOException
    {
      byte[] buffy = new byte[1024];
      ZipOutputStream os = new ZipOutputStream(new FileOutputStream(file));
      for (String name : storage.getElementNames())
      {
        ZipEntry entry = new ZipEntry(name);

        os.putNextEntry(entry);

        InputStream is = storage.getInputStream(name);
        for (int read = 0; (read = is.read(buffy)) > 0;)
        {
          os.write(buffy, 0, read);
        }

        os.closeEntry();
      }
      os.close();
    }
  }

  /**
   * Testmethode
   * 
   * @author Christoph Lutz (D-III-ITD-D101)
   */
  public static void main(String[] args)
  {
    System.out.println("Starting Merge");
    long startTime = System.currentTimeMillis();

    ODFMerger merger = new ODFMerger();
    for (int i = 0; i < 150; ++i)
    {
      try
      {
        merger.add(new ZipedODFFileStorage(new File(
          "testdata/odfmerge/slv_seriendruck/test" + i + ".odt")));
      }
      catch (Exception e)
      {
        e.printStackTrace();
      }
      if (i % 10 == 0) System.out.print(" ");
      if (i % 50 == 0) System.out.print("\n");
      System.out.print("X");
      System.out.flush();
    }

    Storage result = merger.getResultStorage();
    try
    {
      new StorageZipOutput(result).writeToFile(new File("/tmp/test.odt"));
    }
    catch (Exception e)
    {
      e.printStackTrace();
    }

    System.out.println("\n\nMerge finished after "
      + (System.currentTimeMillis() - startTime) + " ms");

  }
}
