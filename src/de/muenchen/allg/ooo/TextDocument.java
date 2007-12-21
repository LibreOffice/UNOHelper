/*
* Dateiname: TextDocument.java
* Projekt  : UNOHelper
* Funktion : Hilfsfunktionen für die Arbeit mit OOo TextDokumenten
* 
* Copyright: Landeshauptstadt München
*
* Änderungshistorie:
* Datum      | Wer | Änderungsgrund
* -------------------------------------------------------------------
* 19.12.2007 | BNK | Erstellung
* -------------------------------------------------------------------
*
* @author Matthias Benkmann (D-III-ITD 5.1)
* @version 1.0
* 
*/
package de.muenchen.allg.ooo;

import java.io.StringReader;
import java.util.Arrays;
import java.util.Collection;
import java.util.Iterator;
import java.util.SortedSet;
import java.util.TreeSet;
import java.util.Vector;
import java.util.regex.Matcher;

import com.sun.star.awt.XControlModel;
import com.sun.star.beans.Property;
import com.sun.star.beans.XPropertySet;
import com.sun.star.beans.XPropertySetInfo;
import com.sun.star.container.XEnumeration;
import com.sun.star.container.XEnumerationAccess;
import com.sun.star.container.XNameContainer;
import com.sun.star.container.XNamed;
import com.sun.star.drawing.XControlShape;
import com.sun.star.lang.XServiceInfo;
import com.sun.star.text.XDependentTextField;
import com.sun.star.text.XText;
import com.sun.star.text.XTextDocument;
import com.sun.star.text.XTextRange;
import com.sun.star.text.XTextTable;
import com.sun.star.uno.UnoRuntime;

import de.muenchen.allg.afid.UNO;
import de.muenchen.allg.itd51.parser.ConfigThingy;
import de.muenchen.allg.itd51.wollmux.Bookmark;
import de.muenchen.allg.itd51.wollmux.Logger;
import de.muenchen.allg.itd51.wollmux.former.DocumentTree.CheckboxNode;
import de.muenchen.allg.itd51.wollmux.former.DocumentTree.ContainerNode;
import de.muenchen.allg.itd51.wollmux.former.DocumentTree.DropdownNode;
import de.muenchen.allg.itd51.wollmux.former.DocumentTree.GroupBookmarkNode;
import de.muenchen.allg.itd51.wollmux.former.DocumentTree.InputNode;
import de.muenchen.allg.itd51.wollmux.former.DocumentTree.InsertionBookmarkNode;
import de.muenchen.allg.itd51.wollmux.former.DocumentTree.ParagraphNode;
import de.muenchen.allg.itd51.wollmux.former.DocumentTree.TextRangeNode;
import de.muenchen.allg.itd51.wollmux.former.insertion.InsertionModel;

public class TextDocument
{
  private static final String[] PAGE_STYLE_PROP_NAMES = { "FollowStyle",
      "BackColor", "BackGraphicURL", "BackGraphicLocation",
      "BackGraphicFilter", "IsLandscape", "PageStyleLayout", "BackTransparent",
      "Size", "Width", "Height", "LeftMargin", "RightMargin", "TopMargin",
      "BottomMargin", "LeftBorder", "RightBorder", "TopBorder", "BottomBorder",
      "LeftBorderDistance", "RightBorderDistance", "TopBorderDistance",
      "BottomBorderDistance", "ShadowFormat", "NumberingType",
      "PrinterPaperTray", "RegisterModeActive", "RegisterParagraphStyle",
      "HeaderIsOn", "HeaderIsShared", "HeaderBackTransparent",
      "HeaderBackColor", "HeaderBackGraphicURL", "HeaderBackGraphicFilter",
      "HeaderBackGraphicLocation", "HeaderLeftMargin", "HeaderRightMargin",
      "HeaderLeftBorder", "HeaderRightBorder", "HeaderTopBorder",
      "HeaderBottomBorder", "HeaderLeftBorderDistance",
      "HeaderRightBorderDistance", "HeaderTopBorderDistance",
      "HeaderBottomBorderDistance", "HeaderShadowFormat", "HeaderBodyDistance",
      "HeaderHeight", "HeaderIsDynamicHeight", "FooterIsOn", "FooterIsShared",
      "FooterBackTransparent", "FooterBackColor", "FooterBackGraphicURL",
      "FooterBackGraphicFilter", "FooterBackGraphicLocation",
      "FooterLeftMargin", "FooterRightMargin", "FooterLeftBorder",
      "FooterRightBorder", "FooterTopBorder", "FooterBottomBorder",
      "FooterLeftBorderDistance", "FooterRightBorderDistance",
      "FooterTopBorderDistance", "FooterBottomBorderDistance",
      "FooterShadowFormat", "FooterBodyDistance", "FooterHeight",
      "FooterIsDynamicHeight", "FootnoteHeight", "FootnoteLineWeight",
      "FootnoteLineColor", "FootnoteLineRelativeWidth", "FootnoteLineAdjust",
      "FootnoteLineTextDistance", "FootnoteLineDistance", "WritingMode",
      "GridDisplay", "GridMode", "GridColor", "GridLines", "GridBaseHeight",
      "GridRubyHeight", "GridRubyBelow", "GridPrint", "HeaderDynamicSpacing",
      "FooterDynamicSpacing", "BorderDistance", "FooterBorderDistance",
      "HeaderBorderDistance" };
  
  /*
   * Fehlen: TextColumns, UserDefinedAttributes, 
   * HeaderText, HeaderTextLeft, HeaderTextRight,  
   * FooterText, FooterTextLeft, FooterTextRight,
   * */
  
  
  /**
   * Kopiert die Eigenschaften von PageStyle oldStyle auf einen PageStyle mit Namen 
   * newName (der angelegt wird, wenn er noch nicht existiert). newName darf sich nicht auf
   * den selben PageStyle beziehen. 
   * @param doc das Textdokument in dem der neue PageStyle angelegt werden soll. Der alte
   * PageStyle kann in einem beliebigen Dokument liegen.
   * @author Matthias Benkmann (D-III-ITD 5.1)
   * TESTED
   * @throws Exception falls was schief geht 
   */
  public static void copyPageStyle(XTextDocument doc, XPropertySet oldStyle, String newName) throws Exception
  {
    XNameContainer pageStyles = UNO.XNameContainer(UNO.XStyleFamiliesSupplier(doc).getStyleFamilies().getByName("PageStyles"));
    
    XPropertySet newStyle;
    if (pageStyles.hasByName(newName))
    {
      newStyle = UNO.XPropertySet(pageStyles.getByName(newName));
    }
    else
    {
      newStyle = UNO.XPropertySet(UNO.XMultiServiceFactory(doc).createInstance("com.sun.star.style.PageStyle"));
      pageStyles.insertByName(newName, newStyle);
    }
        
    for (int i = 0; i < PAGE_STYLE_PROP_NAMES.length; ++i)
    {
      Object val = null;
      try{
        val = null;
        val = oldStyle.getPropertyValue(PAGE_STYLE_PROP_NAMES[i]);
        newStyle.setPropertyValue(PAGE_STYLE_PROP_NAMES[i], val);
      }catch(Exception x)
      {
        if (val != null) //Nur dann Exception werfen, wenn überhaupt ein Property zu kopieren da ist
        {
          throw new Exception("Fehler beim Kopieren von Property \""+PAGE_STYLE_PROP_NAMES[i]+"\"", x);
        }
      }
    }
  }
  
  /**
   * Hängt den Inhalt von source an dest an.
   * @param doc das Dokument in dem sich dest befindet. enu muss *nicht* in doc sein.
   * @author Matthias Benkmann (D-III-ITD 5.1)
   * TODO Testen
   */
  public static void appendXText2XText(XText source, XText dest, XTextDocument doc)
  {
    XEnumerationAccess enuAccess = UNO.XEnumerationAccess(source);
    if (enuAccess == null) return;
    XEnumeration enu = enuAccess.createEnumeration();
    copyParagraphEnumeration(enu, dest, doc);
  }
  
  /**
   * Nimmt eine XEnumeration enu von Absätzen und TextTables und 
   * hängt eine Kopie jedes Elements an dest an.
   * @param doc das Dokument in dem dest liegt.
   * @author Matthias Benkmann (D-III-ITD 5.1)
   */
  private static void copyParagraphEnumeration(XEnumeration enu, XText dest, XTextDocument doc)
  {
    XEnumerationAccess enuAccess;
    while (enu.hasMoreElements())
    {
      Object ele;
      try{ ele = enu.nextElement(); } catch(Exception x){continue;}
      enuAccess = UNO.XEnumerationAccess(ele);
      if (enuAccess != null) //ist wohl ein SwXParagraph
      {
        copyParagraph(enuAccess, dest, doc);
      }
      else //unterstützt nicht XEnumerationAccess, ist wohl SwXTextTable 
      {
        XTextTable table = UNO.XTextTable(ele);
        if (table != null) copyTextTable(table, dest, doc);
      }
    }
  }
  
  /**
   * Nimmt die Inhalte von paragraph und hängt sie an dest an.
   * @param doc das Dokument in dem sich dest befindet. paragraph muss *nicht* in doc sein.
   * @author Matthias Benkmann (D-III-ITD 5.1)
   * TODO Testen
   */
  private static void copyParagraph(XEnumerationAccess paragraph, XText dest, XTextDocument doc)
  {
    /*
     * enumeriere alle TextPortions des Paragraphs
     */
    XEnumeration textPortionEnu = paragraph.createEnumeration();
    while (textPortionEnu.hasMoreElements())
    {
      Object textPortion;
      try{ textPortion = textPortionEnu.nextElement(); } catch(Exception x){continue;};
      
      String textPortionType = (String)UNO.getProperty(textPortion, "TextPortionType");
      if (textPortionType.equals("Text"))
      {
        appendTextPortionString(dest, textPortion);
      } else if (textPortionType.equals("TextField"))
      {
        appendTextPortionString(dest, textPortion);
      } else if (textPortionType.equals("TextContent"))
      {
        //derzeit nicht unterstützt (und ich weiß eh ned, was mit diesem Type geliefert wird)
      } else if (textPortionType.equals("Footnote"))
      {
        appendTextPortionString(dest, textPortion);
      } else if (textPortionType.equals("ControlCharacter"))
      {
        appendTextPortionString(dest, textPortion);
      } else if (textPortionType.equals("ReferenceMark"))
      {
        // ReferenceMarks kopieren derzeit nicht unterstützt
      } else if (textPortionType.equals("DocumentIndexMark"))
      {
        // derzeit nicht unterstützt
      } else if (textPortionType.equals("Bookmark"))
      {
        //Bookmarks kopieren derzeit nicht unterstützt
      } else if (textPortionType.equals("Redline"))
      {
        //derzeit nicht unterstützt
      } else if (textPortionType.equals("Ruby"))
      {
        //derzeit nicht unterstützt
      } else if (textPortionType.equals("Frame"))
      {
        // auch Checkboxen und vermutlich andere Shapes
        // derzeit nicht unterstützt
      }
      else continue; //Dies sollte nicht passieren, da oben alle dokumentierten Typen aufgeführt sind
    }
  }
  
  public static void main(String[] args) throws Exception
  {
    /*
     * Kopiert den Standard-PageStyle des aktuellen Vordergrunddokuments auf den PageStyle
     * "foo" und testet, wieviele Eigenschaften erfolgreich kopiert werden konnten.
     */
    
    UNO.init();
    XTextDocument doc = UNO.XTextDocument(UNO.desktop.getCurrentComponent());
    if (doc == null)
    {
      System.err.println("Vordergrundokument ist kein Textdokument");
      System.exit(0);
    }
    
    XNameContainer pageStyles = UNO.XNameContainer(UNO.XStyleFamiliesSupplier(doc).getStyleFamilies().getByName("PageStyles"));
    
    XPropertySet newStyle;
    String newName = "foo";
    String oldName = "Standard";
    if (pageStyles.hasByName(newName))
    {
      newStyle = UNO.XPropertySet(pageStyles.getByName(newName));
    }
    else
    {
      newStyle = UNO.XPropertySet(UNO.XMultiServiceFactory(doc).createInstance("com.sun.star.style.PageStyle"));
      pageStyles.insertByName(newName, newStyle);
    }
    
    XPropertySet oldStyle = UNO.XPropertySet(pageStyles.getByName(oldName));
    SortedSet propNames = new TreeSet();
    
    XPropertySetInfo oldStylePropInfo = oldStyle.getPropertySetInfo();
    XPropertySetInfo newStylePropInfo = newStyle.getPropertySetInfo();
    
    Property[] props = oldStylePropInfo.getProperties();
    for (int i = 0; i < props.length; ++i)
      propNames.add(props[i].Name);
    
    props = newStylePropInfo.getProperties();
    for (int i = 0; i < props.length; ++i)
      propNames.add(props[i].Name);
    
    // -1 ungesetzt, 0 nach kopieren gleich, 1 nach kopieren ungleich,  
    // 2 vor kopieren gleich
    int[] propCompareState = new int[propNames.size()];
    Arrays.fill(propCompareState, -1);
    
    Iterator iter = propNames.iterator();
    int idx = 0;
    while (iter.hasNext())
    {
      String propName = (String)iter.next();
      if (newStylePropInfo.hasPropertyByName(propName) 
       && oldStylePropInfo.hasPropertyByName(propName))
      {
        if (UnoRuntime.areSame(newStyle.getPropertyValue(propName), oldStyle.getPropertyValue(propName)))
          propCompareState[idx] = 2; 
      }
      
      ++idx;
    }
    
    copyPageStyle(doc, oldStyle, newName);
    
    iter = propNames.iterator();
    idx = 0;
    while (iter.hasNext())
    {
      String propName = (String)iter.next();
      if (!newStylePropInfo.hasPropertyByName(propName))
      {
        System.out.println("*** "+propName+" => nur im alten Style");
      }
      else if (!oldStylePropInfo.hasPropertyByName(propName))
      {
        System.out.println("*** "+propName+" => nur im neuen Style");
      }
      else if (!UnoRuntime.areSame(newStyle.getPropertyValue(propName), oldStyle.getPropertyValue(propName)))
      {
        System.out.println("*** "+propName+" => verschieden!");
        if (propCompareState[idx] == 2)
          System.out.println(propName+" => vor Kopieren noch gleich!!!!!");
      }
      else if (propCompareState[idx] == 2)
        System.out.println(propName+" => schon vor Kopieren gleich!");
      
      
      ++idx;
    }
   
    System.exit(0);
    
  }
  
}
