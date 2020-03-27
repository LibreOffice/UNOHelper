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

import java.util.ArrayList;
import java.util.HashSet;
import java.util.List;
import java.util.Set;
import java.util.regex.Pattern;

import com.sun.star.container.NoSuchElementException;
import com.sun.star.container.XEnumerationAccess;
import com.sun.star.container.XNamed;
import com.sun.star.text.XTextContent;
import com.sun.star.text.XTextRange;

import de.muenchen.allg.afid.UNO;
import de.muenchen.allg.afid.UnoHelperException;
import de.muenchen.allg.afid.UnoHelperRuntimeException;
import de.muenchen.allg.afid.UnoIterator;
import de.muenchen.allg.util.UnoProperty;

/**
 * Hilfsfunktionen für die Arbeit mit OOo TextDokumenten
 *
 * @author Matthias Benkmann (D-III-ITD 5.1)
 */
public class TextDocument
{
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
   * Liefert die Namen aller Bookmarks, die in im Bereich range existieren und regex matchen.
   * Hinweis: Die Funktion liefert nur Bookmarks zurück, von denen Anfang UND Ende in dem Bereich
   * liegen (die schließt kollabierte Bookmarks mit ein). OpenOffice.org hatte in diesem Bereich
   * einige Probleme (oder hat sie immer noch), so dass ich nicht sicher bin, ob dies überall
   * gewährleistet ist. Insbesondere bei TextRanges in Tabellen meine ich mich erinnern zu können,
   * dass die Enumeration immer die ganze Zelle liefert anstatt den ausgewählten Bereich. Es ist
   * also nicht auszuschließen, dass diese Methode unter bestimmten Umständen zu viele Bookmarks
   * zurückliefert.
   * 
   * @throws UnoHelperRuntimeException
   */
  public static Set<String> getBookmarkNamesMatching(Pattern regex, XTextRange range)
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
            } else
            {
              started.add(name);
            }
          } else if (started.contains(name))
          {
            found.add(name);
          }
        }
      } catch (UnoHelperException e)
      {
        throw new UnoHelperRuntimeException(e);
      }
    });

    return found;
  }

  /**
   * Gibt eine Liste aller Bookmarks in einer TextRange zurück.
   *
   * @throws UnoHelperRuntimeException
   */
  public static List<XNamed> getBookmarkByTextRange(XTextRange range)
  {
    ArrayList<XNamed> bookmarks = new ArrayList<>();

    UNO.forEachTextPortionInRange(range, o -> {
      try
      {
        bookmarks.add(UNO.XNamed(UnoProperty.getProperty(o, UnoProperty.BOOKMARK)));
      } catch (UnoHelperException e)
      {
        throw new UnoHelperRuntimeException(e);
      }
    });

    return bookmarks;
  }
}
