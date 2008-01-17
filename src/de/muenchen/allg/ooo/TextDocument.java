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

import java.util.Arrays;
import java.util.Iterator;
import java.util.SortedSet;
import java.util.TreeSet;

import com.sun.star.beans.Property;
import com.sun.star.beans.XPropertySet;
import com.sun.star.beans.XPropertySetInfo;
import com.sun.star.container.XNameContainer;
import com.sun.star.text.XAutoTextContainer;
import com.sun.star.text.XAutoTextEntry;
import com.sun.star.text.XAutoTextGroup;
import com.sun.star.text.XText;
import com.sun.star.text.XTextDocument;
import com.sun.star.text.XTextRange;
import com.sun.star.uno.UnoRuntime;

import de.muenchen.allg.afid.UNO;

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
      "HeaderBorderDistance" }; /* Fehlen: TextColumns, UserDefinedAttributes 
   * HeaderText, HeaderTextLeft, HeaderTextRight,  
   * FooterText, FooterTextLeft, FooterTextRight,
   * */
  
  private static final String[] HEADER_FOOTER_PROP_NAMES = { "HeaderText", "HeaderTextLeft",
    "HeaderTextRight", "FooterText", "FooterTextLeft", "FooterTextRight"};
  
  /**
   * Kopiert die Eigenschaften von PageStyle oldStyle auf einen PageStyle mit Namen 
   * newName (der angelegt wird, wenn er noch nicht existiert).
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
      if (UnoRuntime.areSame(oldStyle, newStyle)) return;
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
    
    for (int i = 0; i < HEADER_FOOTER_PROP_NAMES.length; ++i)
    {
      XText val = null;
      try{
        val = null;
        val = UNO.XText(oldStyle.getPropertyValue(HEADER_FOOTER_PROP_NAMES[i]));
        if (val != null)
        {
          XText dest = UNO.XText(newStyle.getPropertyValue(HEADER_FOOTER_PROP_NAMES[i]));
          copyXText2XTextRange(val, dest.createTextCursorByRange(dest));
        }
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
   * Ersetzt dest durch den Inhalt von source.
   * @author Matthias Benkmann (D-III-ITD 5.1)
   * TODO Testen
   * @throws Exception falls was schief geht 
   */
  public static void copyXText2XTextRange(XText source, XTextRange dest) throws Exception
  {
    XAutoTextContainer autoTextContainer = UNO.XAutoTextContainer(UNO.createUNOService("com.sun.star.text.AutoTextContainer"));
    int rand;
    
    XAutoTextGroup atgroup;
    for (int i = 0; ; ++i)
    {
      rand = (int)(Math.random()*100000);
      try{
        atgroup = autoTextContainer.insertNewByName("WollMuxTemp"+rand+"*1");
        break;
      }catch(Exception x)  
      {
        if ( i >= 100) throw x;
      }
    }
    
    try{
      XTextRange range = source.createTextCursorByRange(source);
      XAutoTextEntry atentry = atgroup.insertNewByName("X","X",range);
      dest.setString("");
      atentry.applyTo(dest);
    }
    finally{
      try{ autoTextContainer.removeByName("WollmuxTemp"+rand+"*1"); }catch(Exception y){};
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
