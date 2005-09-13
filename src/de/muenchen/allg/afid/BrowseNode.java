/*
 * Hilfen für die Arbeit mit BrowseNodes
* Dateiname: BrowseNode.java
* Projekt  : n/a
* Funktion : Hilfen für die Arbeit mit BrowseNodes
* 
* Copyright: Landeshauptstadt München
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
import java.util.NoSuchElementException;
import java.util.Vector;

import com.sun.star.script.browse.XBrowseNode;
import com.sun.star.uno.UnoRuntime;

/**
 * Diese Klasse vereinfacht die Arbeit mit Objekten, die den Dienst BrowseNode
 * unterstützen. Diese sind vor allem
 * nützlich beim Durchsuchen des Baumes aller installierten Skripts.
 * @author bnk
 */
public class BrowseNode 
{
	private XBrowseNode node;
	public BrowseNode(XBrowseNode node) {this.node = node;}
	
	/**
	 * Gibt die URL des Makros zurück, wie sie für den Aufruf über 
	 * das Skripting-Framework benötigt wird. Nicht zu verwechseln mit "macro:" URLs!
	 * Falls dieses Property nicht existiert, wird null geliefert.
	 * @author bnk
	 */
	public String URL() {return (String)UNO.getProperty(node, "URI");}
	
	/**
	 * Liefert das ungewrappte Objekt für die direkte Übergabe an UNO-Funktionen.
	 * @author bnk
	 */
	public XBrowseNode unwrap() {return node;}
	
	/**
	 * Liefert den Namen des Knoten.
	 * @author bnk
	 */
	public String getName() {return node.getName(); }
	
	/**
	 * Liefert den Typ des Knoten. Im Zusammenhang mit Makros sind die möglichen 
	 * Werte aus der Konstantengruppe {@link com.sun.star.script.browse.BrowseNodeTypes}.
	 * @author bnk
	 */
	public short getType() {return node.getType();}
	
	/**
	 * Abkürzung für queryInterface(). 
	 * @param c spezifiziert das Interface das gequeryt werden soll.
	 * @author bnk
	 */
	public Object as(Class c) 
	{
		return UnoRuntime.queryInterface(c, node);
	}
	
	/** Dieser Iterator liefert den Knoten und alle Abkömmlinge als 
	 * {@link BrowseNode}s.*/
	public Iterator iterator() {return new ChildIterator(node, false);};
	
	/** Dieser Iterator liefert alle Kinder des Knoten als {@link BrowseNode}s.*/
	public Iterator children() {return new ChildIterator(node, true);};
	
	protected static class ChildIterator implements Iterator 
	{
		private boolean childrenOnly;
	  private Vector toVisit = new Vector();
		
		public ChildIterator(XBrowseNode root, boolean childrenOnly)
		{
			this.childrenOnly = childrenOnly;
		  toVisit.add(root);
		  if (childrenOnly) expandLast();
		}
		
		/** @throws UnsupportedOperationException*/
		public void remove() {throw new UnsupportedOperationException();}
		public boolean hasNext() { return !toVisit.isEmpty() ;}
		
		/** Returns an {@link Object} of class {@BrowseNode}. */
		public Object next() throws NoSuchElementException
		{
			XBrowseNode retval = (XBrowseNode)toVisit.lastElement(); 
			if (childrenOnly)
			  toVisit.remove(toVisit.size()-1);
			else
			  expandLast();
			
			return new BrowseNode(retval); 
		}
		
		protected void expandLast()
		{
			if (toVisit.isEmpty()) throw new NoSuchElementException();
			
			/* Die try-catch Blöcke sind zum Schutz gegen Bugs im SFW. Ich hatte z.B.
			 * einen Bug im Python-Modul, der eine Exception geworfen hat beim Aufruf
			 * von hasChildNodes(). */
			try{
				XBrowseNode xBrowseNode = (XBrowseNode)toVisit.remove(toVisit.size()-1);
				
				if (xBrowseNode.hasChildNodes())
				{
					XBrowseNode[] child = xBrowseNode.getChildNodes();
					for (int i = child.length - 1; i >= 0; --i)
						toVisit.add(child[i]);	  	
				}
			}catch(Exception e){}
		}
	}
}
