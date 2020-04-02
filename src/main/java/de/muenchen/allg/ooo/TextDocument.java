/*
* Dateiname: TextDocument.java
* Projekt  : UNOHelper
* Funktion : Hilfsfunktionen für die Arbeit mit OOo TextDokumenten
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
* 19.12.2007 | BNK | Erstellung
* 15.01.2008 | BNK | +copyDirectValueCharAttributes
* 29.01.2008 | BNK | +copySimpleProperties()
* 29.01.2008 | BNK | +deleteParagraph
* 30.01.2008 | BNK | +disappearParagraph
* 03.03.2010 | ERT | +getBookmarkNamesStartingWith
* 17.03.2011 | BED | isSimpleType um Prüfung auf Void und Type erweitert
* -------------------------------------------------------------------
*
* @author Matthias Benkmann (D-III-ITD 5.1)
* 
*/
package de.muenchen.allg.ooo;

import java.lang.reflect.Array;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.HashSet;
import java.util.Iterator;
import java.util.List;
import java.util.Set;
import java.util.SortedSet;
import java.util.TreeSet;
import java.util.regex.Pattern;

import com.sun.star.beans.Property;
import com.sun.star.beans.PropertyState;
import com.sun.star.beans.PropertyVetoException;
import com.sun.star.beans.UnknownPropertyException;
import com.sun.star.beans.XPropertySet;
import com.sun.star.beans.XPropertySetInfo;
import com.sun.star.beans.XPropertyState;
import com.sun.star.container.NoSuchElementException;
import com.sun.star.container.XEnumeration;
import com.sun.star.container.XEnumerationAccess;
import com.sun.star.container.XNameContainer;
import com.sun.star.container.XNamed;
import com.sun.star.lang.IllegalArgumentException;
import com.sun.star.lang.WrappedTargetException;
import com.sun.star.style.LineSpacing;
import com.sun.star.style.LineSpacingMode;
import com.sun.star.text.XAutoTextContainer;
import com.sun.star.text.XAutoTextEntry;
import com.sun.star.text.XAutoTextGroup;
import com.sun.star.text.XText;
import com.sun.star.text.XTextContent;
import com.sun.star.text.XTextDocument;
import com.sun.star.text.XTextRange;
import com.sun.star.uno.AnyConverter;
import com.sun.star.uno.UnoRuntime;

import de.muenchen.allg.afid.UNO;
import de.muenchen.allg.afid.UnoHelperException;
import de.muenchen.allg.afid.UnoHelperRuntimeException;
import de.muenchen.allg.afid.UnoIterator;
import de.muenchen.allg.util.UnoComponent;
import de.muenchen.allg.util.UnoProperty;

/**
 * Hilfsfunktionen für die Arbeit mit OOo TextDokumenten
 *
 * @author Matthias Benkmann (D-III-ITD 5.1)
 */
public class TextDocument
{
  /**
   * Diese Properties werden in der hier gelisteten Reihenfolge über
   * getProperty/setProperty Paare kopiert.
   */
  private static final String[] PAGE_STYLE_PROP_NAMES = { "BackColor",
      "BackGraphicURL", "BackGraphicLocation", "BackGraphicFilter",
      "IsLandscape", "PageStyleLayout", "BackTransparent", "Size", "Width",
      "Height", "LeftMargin", "RightMargin", "TopMargin", "BottomMargin",
      "LeftBorder", "RightBorder", "TopBorder", "BottomBorder",
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
      "HeaderBorderDistance" }; /*
                                 * Fehlen: TextColumns, UserDefinedAttributes
                                 * HeaderText, HeaderTextLeft, HeaderTextRight,
                                 * FooterText, FooterTextLeft, FooterTextRight,
                                 * FollowStyle (weil es zu Absturz führt (siehe
                                 * CrashCopyingPagestyle.java)
                                 */

  private static final String[] CHAR_STYLE_PROP_NAMES = { "CharFontName",
      "CharFontStyleName", "CharFontFamily", "CharFontCharSet", "CharFontPitch",
      "CharColor", "CharEscapement", "CharHeight", "CharUnderline",
      "CharWeight", "CharPosture", "CharAutoKerning", "CharBackColor",
      "CharBackTransparent", "CharCaseMap", "CharCrossedOut", "CharFlash",
      "CharStrikeout", "CharWordMode", "CharKerning", "CharLocale",
      "CharKeepTogether", "CharNoLineBreak", "CharShadowed", "CharFontType",
      "CharStyleName", "CharContoured", "CharCombineIsOn", "CharCombinePrefix",
      "CharCombineSuffix", "CharEmphasis", "CharRelief", "RubyText",
      "RubyAdjust", "RubyCharStyleName", "RubyIsAbove", "CharRotation",
      "CharRotationIsFitToLine", "CharScaleWidth", "HyperLinkURL",
      "HyperLinkTarget", "HyperLinkName", "VisitedCharStyleName",
      "UnvisitedCharStyleName", "CharEscapementHeight", "CharNoHyphenation",
      "CharUnderlineColor", "CharUnderlineHasColor", "CharHidden",
      "CharHeightAsian", "CharWeightAsian", "CharFontNameAsian",
      "CharFontStyleNameAsian", "CharFontFamilyAsian", "CharFontCharSetAsian",
      "CharFontPitchAsian", "CharPostureAsian", "CharLocaleAsian",
      "CharHeightComplex", "CharWeightComplex", "CharFontNameComplex",
      "CharFontStyleNameComplex", "CharFontFamilyComplex",
      "CharFontCharSetComplex", "CharFontPitchComplex", "CharPostureComplex",
      "CharLocaleComplex" };/* Fehlen: TextUserDefinedAttributes */

  /**
   * Diese Properties werden in der hier gelisteten Reihenfolge über
   * {@link #copyXText2XTextRange(XText, XTextRange)} kopiert.
   */
  private static final String[] HEADER_FOOTER_PROP_NAMES = { "HeaderText",
      "HeaderTextLeft", "HeaderTextRight", "FooterText", "FooterTextLeft",
      "FooterTextRight" };

  /**
   * Kopiert die Eigenschaften von PageStyle oldStyle auf einen PageStyle mit
   * Namen newName (der angelegt wird, wenn er noch nicht existiert).
   * 
   * @param doc
   *          das Textdokument in dem der neue PageStyle angelegt werden soll.
   *          Der alte PageStyle kann in einem beliebigen Dokument liegen.
   * @throws Exception
   *           falls was schief geht
   */
  public static void copyPageStyle(XTextDocument doc, XPropertySet oldStyle,
      String newName) throws Exception
  {
    XNameContainer pageStyles = UNO
        .XNameContainer(UNO.XStyleFamiliesSupplier(doc).getStyleFamilies()
            .getByName("PageStyles"));

    XPropertySet newStyle;
    if (pageStyles.hasByName(newName))
    {
      newStyle = UNO.XPropertySet(pageStyles.getByName(newName));
      if (UnoRuntime.areSame(oldStyle, newStyle))
        return;
    } else
    {
      newStyle = UNO.XPropertySet(UNO.XMultiServiceFactory(doc)
          .createInstance("com.sun.star.style.PageStyle"));
      pageStyles.insertByName(newName, newStyle);
    }

    for (int i = 0; i < PAGE_STYLE_PROP_NAMES.length; ++i)
    {
      Object val = null;
      try
      {
        val = null;
        val = oldStyle.getPropertyValue(PAGE_STYLE_PROP_NAMES[i]);
        newStyle.setPropertyValue(PAGE_STYLE_PROP_NAMES[i], val);
      } catch (Exception x)
      {
        if (val != null) // Nur dann Exception werfen, wenn überhaupt ein
                         // Property zu kopieren da ist
        {
          throw new Exception("Fehler beim Kopieren von Property \""
              + PAGE_STYLE_PROP_NAMES[i] + "\"", x);
        }
      }
    }

    for (int i = 0; i < HEADER_FOOTER_PROP_NAMES.length; ++i)
    {
      XText val = null;
      try
      {
        val = UNO.XText(oldStyle.getPropertyValue(HEADER_FOOTER_PROP_NAMES[i]));
        if (val != null)
        {
          XText dest = UNO
              .XText(newStyle.getPropertyValue(HEADER_FOOTER_PROP_NAMES[i]));
          copyXText2XTextRange(val, dest.createTextCursorByRange(dest));
        }
      } catch (Exception x)
      {
        throw new Exception("Fehler beim Kopieren von Property \""
            + HEADER_FOOTER_PROP_NAMES[i] + "\"", x);
      }
    }
  }

  /**
   * Ersetzt dest durch den Inhalt von source.
   *
   * @throws Exception
   *           falls was schief geht
   */
  public static void copyXText2XTextRange(XText source, XTextRange dest)
      throws Exception
  {
    XAutoTextContainer autoTextContainer = UNO.XAutoTextContainer(
        UnoComponent.createComponentWithContext(UnoComponent.CSS_TEXT_AUTO_TEXT_CONTAINER));
    int rand;

    String groupName = null;
    XAutoTextGroup atgroup;
    for (int i = 0;; ++i)
    {
      rand = (int) (Math.random() * 100000);
      try
      {
        groupName = "WollMuxTemp" + rand;
        atgroup = autoTextContainer.insertNewByName(groupName + "*1");
        break;
      } catch (Exception x)
      {
        if (i >= 100)
          throw x;
      }
    }

    try
    {
      XTextRange range = source.createTextCursorByRange(source);
      XAutoTextEntry atentry = atgroup.insertNewByName("X", "X", range);
      dest.setString("");
      atentry.applyTo(dest);
    } finally
    {
      try
      {
        autoTextContainer.removeByName(groupName);
      } catch (Exception y)
      {
      }
    }
  }

  /**
   * Kopiert alle Character Attributes, die den Zustand
   * {@link com.sun.star.beans.PropertyState#DIRECT_VALUE} haben von from nach
   * to.
   */
  public static void copyDirectValueCharAttributes(XPropertyState from,
      XPropertySet to)
  {
    XPropertySet fromSet = UNO.XPropertySet(from);
    for (int i = 0; i < CHAR_STYLE_PROP_NAMES.length; ++i)
    {
      try
      {
        if (from.getPropertyState(CHAR_STYLE_PROP_NAMES[i])
            .equals(PropertyState.DIRECT_VALUE))
        {
          to.setPropertyValue(CHAR_STYLE_PROP_NAMES[i],
              fromSet.getPropertyValue(CHAR_STYLE_PROP_NAMES[i]));
        }
      } catch (Exception x)
      {
      }
    }
  }

  /**
   * Löscht den ganzen ersten Absatz an der Cursorposition textCursor.
   */
  public static void deleteParagraph(XTextRange range)
  {
    // Beim Löschen des Absatzes erzeugt OOo ein ungewolltes
    // "Zombie"-Bookmark.
    // Issue Siehe http://qa.openoffice.org/issues/show_bug.cgi?id=65247

    XTextContent paragraph = null;

    // Ersten Absatz des Bookmarks holen:
    XEnumerationAccess access = UNO.XEnumerationAccess(range);
    if (access != null)
    {
    	paragraph = new UnoIterator<XTextContent>(access, XTextContent.class).next();
    }

    if (paragraph != null)
    {
	    // Lösche den Paragraph
	    try
	    {
	      // Ist der Paragraph der einzige Paragraph des Textes, dann kann er mit
	      // removeTextContent nicht gelöscht werden. In diesme Fall wird hier
	      // wenigstens der Inhalt entfernt:
	      paragraph.getAnchor().setString("");
	
	      // Paragraph löschen
	      range.getText().removeTextContent(paragraph);
	    } catch (NoSuchElementException e)
	    {
	      // sollte eigentlich nicht passieren können
	    }
    }
  }

  /**
   * Lässt den ersten Absatz in der TextRange range verschwinden, ohne ihn zu
   * löschen.
   */
  public static void disappearParagraph(XTextRange range)
  {
    try
    {
      XEnumerationAccess access = UNO.XEnumerationAccess(range);
      XEnumeration xenum = access.createEnumeration();
      XPropertySet props = UNO.XPropertySet(xenum.nextElement());
      props.setPropertyValue("CharHeight", Float.valueOf(0.1f));
      props.setPropertyValue("CharHeightAsian", Float.valueOf(0.1f));
      props.setPropertyValue("CharHeightComplex", Float.valueOf(0.1f));
      props.setPropertyValue("CharHidden", Boolean.TRUE);
      LineSpacing ls = new LineSpacing();
      ls.Mode = LineSpacingMode.FIX;
      ls.Height = 2; // Lieber nicht 0 nehmen. Vielleicht möchte jemand irgendwo
                     // dadurch dividieren
      props.setPropertyValue("ParaLineSpacing", ls);
      props.setPropertyValue("ParaTopMargin", Integer.valueOf(0));
      props.setPropertyValue("ParaBottomMargin", Integer.valueOf(0));
      props.setPropertyValue("ParaLineNumberCount", Boolean.FALSE);
    } catch (Exception x)
    {
      // sollte nicht passieren
    }
  }

  /**
   * Kopiert alle Properties, die nicht komplexe Objekte (structs, interfaces,
   * exceptions) oder Sequenzen von komplexen Objekten sind von in nach out.
   * Sequenzen simpler Typen werden kopiert.
   */
  public static void copySimpleProperties(XPropertySet in, XPropertySet out)
  {
    XPropertySetInfo propInfo = in.getPropertySetInfo();
    Property[] props = propInfo.getProperties();
    for (Property prop : props)
    {
      Object value;
      try
      {
        value = in.getPropertyValue(prop.Name);
        if (isSimpleType(value))
        {
            out.setPropertyValue(prop.Name, value);
        }
      } catch (UnknownPropertyException | WrappedTargetException | 
      		IllegalArgumentException | PropertyVetoException ex)  {}
    }
  }

  /**
   * Liefert true gdw o nicht Exception, Struct oder Interface ist oder eine
   * Sequenz, die sowas enthält.
   */
  private static boolean isSimpleType(Object o)
  {
    if (AnyConverter.isArray(o))
    {
      try
      {
        Object arryConv = AnyConverter.toArray(o);
        int len = Array.getLength(arryConv);
        if (len == 0)
          return true;
        return isSimpleType(Array.get(arryConv, 0));
      } catch (java.lang.IllegalArgumentException x)
      {
        throw new RuntimeException("Dies dürfte nicht passieren.", x);
      }
    } else if (AnyConverter.isEnum(o))
    {
      return true;
    } else if (AnyConverter.isObject(o))
    {
      return false;
    } else if (AnyConverter.isVoid(o))
    {
      return false;
    } else if (AnyConverter.isType(o))
    {
      return false;
    }

    return true;
  }

  public static void main(String[] args) throws Exception
  {
    /*
     * Kopiert den Standard-PageStyle des aktuellen Vordergrunddokuments auf den
     * PageStyle "foo" und testet, wieviele Eigenschaften erfolgreich kopiert
     * werden konnten.
     */

    UNO.init();
    XTextDocument doc = UNO.XTextDocument(UNO.desktop.getCurrentComponent());
    if (doc == null)
    {
      System.err.println("Vordergrundokument ist kein Textdokument");
      System.exit(0);
    }

    XNameContainer pageStyles = UNO
        .XNameContainer(UNO.XStyleFamiliesSupplier(doc).getStyleFamilies()
            .getByName("PageStyles"));

    XPropertySet newStyle;
    String newName = "foo";
    String oldName = "Standard";
    if (pageStyles.hasByName(newName))
    {
      newStyle = UNO.XPropertySet(pageStyles.getByName(newName));
    } else
    {
      newStyle = UNO.XPropertySet(UNO.XMultiServiceFactory(doc)
          .createInstance("com.sun.star.style.PageStyle"));
      pageStyles.insertByName(newName, newStyle);
    }

    XPropertySet oldStyle = UNO.XPropertySet(pageStyles.getByName(oldName));
    SortedSet<String> propNames = new TreeSet<>();

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

    Iterator<String> iter = propNames.iterator();
    int idx = 0;
    while (iter.hasNext())
    {
      String propName = iter.next();
      if (newStylePropInfo.hasPropertyByName(propName)
          && oldStylePropInfo.hasPropertyByName(propName)
          && UnoRuntime.areSame(newStyle.getPropertyValue(propName),
              oldStyle.getPropertyValue(propName)))
      {
        propCompareState[idx] = 2;
      }

      ++idx;
    }

    copyPageStyle(doc, oldStyle, newName);

    iter = propNames.iterator();
    idx = 0;
    while (iter.hasNext())
    {
      String propName = iter.next();
      if (!newStylePropInfo.hasPropertyByName(propName))
      {
        System.out.println("*** " + propName + " => nur im alten Style");
      } else if (!oldStylePropInfo.hasPropertyByName(propName))
      {
        System.out.println("*** " + propName + " => nur im neuen Style");
      } else if (!UnoRuntime.areSame(newStyle.getPropertyValue(propName),
          oldStyle.getPropertyValue(propName)))
      {
        System.out.println("*** " + propName + " => verschieden!");
        if (propCompareState[idx] == 2)
          System.out.println(propName + " => vor Kopieren noch gleich!!!!!");
      } else if (propCompareState[idx] == 2)
        System.out.println(propName + " => schon vor Kopieren gleich!");

      ++idx;
    }

    System.exit(0);

  }

  /**
   * Liefert die Namen aller Bookmarks, die in im Bereich range existieren und
   * regex matchen. Hinweis: Die Funktion liefert nur Bookmarks zurück, von
   * denen Anfang UND Ende in dem Bereich liegen (die schließt kollabierte
   * Bookmarks mit ein). OpenOffice.org hatte in diesem Bereich einige Probleme
   * (oder hat sie immer noch), so dass ich nicht sicher bin, ob dies überall
   * gewährleistet ist. Insbesondere bei TextRanges in Tabellen meine ich mich
   * erinnern zu können, dass die Enumeration immer die ganze Zelle liefert
   * anstatt den ausgewählten Bereich. Es ist also nicht auszuschließen, dass
   * diese Methode unter bestimmten Umständen zu viele Bookmarks zurückliefert.
   * 
   * @throws UnoHelperRuntimeException
   */
  public static Set<String> getBookmarkNamesMatching(Pattern regex,
      XTextRange range)
  {
    // Hier findet eine iteration des über den XEnumerationAccess des ranges
    // statt. Man könnte statt dessen auch über range-compare mit den bereits
    // bestehenden Blöcken aus TextDocumentModel.get<blockname>Blocks()
    // vergleichen...
    HashSet<String> found = new HashSet<>();
    HashSet<String> started = new HashSet<>();
    
    UNO.forEachTextPortionInRange(range, o -> {
      try
      {
        XNamed bookmark = UNO.XNamed(UnoProperty.getProperty(o, UnoProperty.BOOKMARK));
        String name = (bookmark != null) ? bookmark.getName() : "";

        if (regex.matcher(name).matches())
        {
          if (Boolean.TRUE.equals(UnoProperty.getProperty(o, UnoProperty.IS_START)))
          {
            if (Boolean.TRUE.equals(UnoProperty.getProperty(o, UnoProperty.IS_COLLAPSED)))
            {
              found.add(name);
            }
            else
            {
              started.add(name);
            }
          } 
          else if (started.contains(name))
          {
            found.add(name);
          }
        }
      }
      catch (UnoHelperException e)
      {
        throw new UnoHelperRuntimeException(e);
      }
    });
    
    return found;
  }
  
  public static String parseHighlightColor(Integer highlightColor)
  {
    if (highlightColor == null)
    {
    	return null;
    }
    
    String colStr = "00000000";
    colStr += Integer.toHexString(highlightColor.intValue());
    colStr = colStr.substring(colStr.length() - 8, colStr.length());
    String hcAtt = " HIGHLIGHT_COLOR '" + colStr + "'";
    
    return hcAtt;
  }

  /**
   * Gibt eine Liste aller Bookmarks in einer TextRange zurück.
   * 
   * @throws UnoHelperRuntimeException 
   */
  public static List<XNamed> getBookmarkByTextRange(XTextRange range) {
    ArrayList<XNamed> bookmarks = new ArrayList<>();
    
    UNO.forEachTextPortionInRange(range, o ->
    {
    	try
      {
        bookmarks.add(UNO.XNamed(UnoProperty.getProperty(o, UnoProperty.BOOKMARK)));
      }
      catch (UnoHelperException e)
      {
        throw new UnoHelperRuntimeException(e);
      }
    });

    return bookmarks;
  }
}
