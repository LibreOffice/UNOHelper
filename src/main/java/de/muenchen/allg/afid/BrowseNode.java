/*
 * Hilfen für die Arbeit mit BrowseNodes
* Dateiname: BrowseNode.java
* Projekt  : n/a
* Funktion : Hilfen für die Arbeit mit BrowseNodes
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
*  Datum     | Wer | Änderungsgrund
* -------------------------------------------------------------------
* 07.07.2005 | BNK | Erstellung
* 16.08.2005 | BNK | korrekte Dienststellenbezeichnung
* 17.08.2005 | BNK | bessere Kommentare
* 31.08.2005 | BNK | +children(), um nur Kinder durchzuiterieren
* 13.09.2005 | BNK | besserer Kommentar zu URL()
* -------------------------------------------------------------------
*
* @author D-III-ITD 5.1 Matthias S. Benkmann
* @version 1.0
* 
*/
package de.muenchen.allg.afid;

import java.util.Iterator;
import java.util.LinkedList;
import java.util.NoSuchElementException;

import com.sun.star.script.browse.XBrowseNode;
import com.sun.star.uno.UnoRuntime;

/**
 * Diese Klasse vereinfacht die Arbeit mit Objekten, die den Dienst BrowseNode
 * unterstützen. Diese sind vor allem nützlich beim Durchsuchen des Baumes aller
 * installierten Skripts.
 * 
 */
public class BrowseNode
{
  private XBrowseNode node;

  public BrowseNode(XBrowseNode node)
  {
    this.node = node;
  }

  /**
   * Gibt die URL des Makros zurück, wie sie für den Aufruf über das
   * Skripting-Framework benötigt wird. Nicht zu verwechseln mit "macro:" URLs!
   * Falls dieses Property nicht existiert, wird null geliefert.
   *
   * @throws UnoHelperException 
   */
  public String getURL() throws UnoHelperException
  {
    return (String) UNO.getProperty(node, "URI");
  }
  
  /**
   * Liefert das ungewrappte Objekt für die direkte Übergabe an UNO-Funktionen.
   */
  public XBrowseNode unwrap()
  {
    return node;
  }

  /**
   * Liefert den Namen des Knoten.
   */
  public String getName()
  {
    return node.getName();
  }

  /**
   * Liefert den Typ des Knoten. Im Zusammenhang mit Makros sind die möglichen
   * Werte aus der Konstantengruppe
   * {@link com.sun.star.script.browse.BrowseNodeTypes}.
   */
  public short getType()
  {
    return node.getType();
  }

  /**
   * Abkürzung für queryInterface().
   *
   * @param c
   *          spezifiziert das Interface das gequeryt werden soll.
   */
  public <T> T as(Class<T> c)
  {
    return UnoRuntime.queryInterface(c, node);
  }

  /**
   * Dieser Iterator liefert den Knoten und alle Abkömmlinge als
   * {@link BrowseNode}s.
   */
  public Iterator<BrowseNode> iterator()
  {
    return new ChildIterator(node, false);
  }

  /** Dieser Iterator liefert alle Kinder des Knoten als {@link BrowseNode}s. */
  public Iterator<BrowseNode> children()
  {
    return new ChildIterator(node, true);
  }

  protected static class ChildIterator implements Iterator<BrowseNode>
  {
    private boolean childrenOnly;
    private LinkedList<XBrowseNode> toVisit = new LinkedList<>();

    public ChildIterator(XBrowseNode root, boolean childrenOnly)
    {
      this.childrenOnly = childrenOnly;
      toVisit.add(root);
      if (childrenOnly)
        expandLast();
    }

    /**
     * @throws UnsupportedOperationException
     */
    @Override
    public void remove()
    {
      throw new UnsupportedOperationException();
    }

    @Override
    public boolean hasNext()
    {
      return !toVisit.isEmpty();
    }

    /** Returns an {@link Object} of class {@link BrowseNode}. */
    @Override
    public BrowseNode next()
    {
      XBrowseNode retval = toVisit.getLast();
      if (childrenOnly)
        toVisit.remove(toVisit.size() - 1);
      else
        expandLast();

      return new BrowseNode(retval);
    }

    protected void expandLast()
    {
      if (toVisit.isEmpty())
        throw new NoSuchElementException();

      /*
       * Die try-catch Blöcke sind zum Schutz gegen Bugs im SFW. Ich hatte z.B.
       * einen Bug im Python-Modul, der eine Exception geworfen hat beim Aufruf
       * von hasChildNodes().
       */
      try
      {
        XBrowseNode xBrowseNode = toVisit.remove(toVisit.size() - 1);

        if (xBrowseNode.hasChildNodes())
        {
          XBrowseNode[] child = xBrowseNode.getChildNodes();
          for (int i = child.length - 1; i >= 0; --i)
            toVisit.add(child[i]);
        }
      } catch (Exception e)
      {
      }
    }
  }
}
