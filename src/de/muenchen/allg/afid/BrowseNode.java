/*
 * Hilfen für die Arbeit mit XBrowseNodes
* Dateiname: BrowseNode.java
* Projekt  : n/a
* Funktion : Hilfen für die Arbeit mit XBrowseNodes
* 
* Copyright: Landeshauptstadt München
*
* Änderungshistorie:
* Nr. |  Datum     |   Autor   | Änderungsgrund
* -------------------------------------------------------------------
* 001 | 07.07.2005 |    BNK    | Erstellung
* 002 | 16.08.2005 |    BNK    | korrekte Dienststellenbezeichnung
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
 * @author bnk
 *
 * TODO To change the template for this generated type comment go to
 * Window - Preferences - Java - Code Style - Code Templates
 */
public class BrowseNode 
{
	private XBrowseNode node;
	public BrowseNode(XBrowseNode node) {this.node = node;}
	
	public String URL() {return (String)UNO.getProperty(node, "URI");}
	public XBrowseNode unwrap() {return node;}
	public String getName() {return node.getName(); }
	public short getType() {return node.getType();}
	public Object as(Class c) 
	{
		return UnoRuntime.queryInterface(c, node);
	}
	
	/** The Iterator returns the node and all descendants as 
	 * {@link BrowseNode} objects.*/
	public Iterator iterator() {return new ChildIterator(node);};
	
	protected static class ChildIterator implements Iterator 
	{
		private Vector toVisit = new Vector();
		
		public ChildIterator(XBrowseNode root)
		{
			toVisit.add(root);
		}
		
		/** @throws UnsupportedOperationException*/
		public void remove() {throw new UnsupportedOperationException();}
		public boolean hasNext() { return !toVisit.isEmpty() ;}
		
		/** Returns an {@link Object} of class {@BrowseNode}. */
		public Object next() throws NoSuchElementException
		{
			XBrowseNode retval = (XBrowseNode)toVisit.lastElement(); 
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
