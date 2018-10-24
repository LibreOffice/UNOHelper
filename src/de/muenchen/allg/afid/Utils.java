package de.muenchen.allg.afid;

import java.util.Iterator;
import java.util.List;
import java.util.ListIterator;

import com.sun.star.script.browse.BrowseNodeTypes;
import com.sun.star.script.browse.XBrowseNode;
import com.sun.star.script.provider.XScript;
import com.sun.star.script.provider.XScriptProvider;
import com.sun.star.uno.RuntimeException;

/**
   * Interne Funktionen
   */
  class Utils
  {

    public static class FindNode extends XBrowseNodeAndXScriptProvider
    {
      public String location;
      public boolean isCaseCorrect;

      public FindNode(XBrowseNode xBrowseNode, XScriptProvider xScriptProvider,
          String location, boolean isCaseCorrect)
      {
        super(xBrowseNode, xScriptProvider);
        this.location = location;
        this.isCaseCorrect = isCaseCorrect;
      }

      /**
       * Returns true if <code>this</code> isCaseCorrect and fn2 is not, or both
       * have the same isCaseCorrect but <code>this.location</code> occurs
       * earlier in locations than <code>fn2.location</code>.
       * 
       * @author bnk
       */
      public boolean betterMatchThan(Object fn2, String[] locations)
      {
        Utils.FindNode f2 = (Utils.FindNode) fn2;
        if (this.isCaseCorrect && !f2.isCaseCorrect)
          return true;
        if (f2.isCaseCorrect && !this.isCaseCorrect)
          return false;
        int i = 0;
        int i2 = 0;
        while (i < locations.length && !locations[i].equals(this.location))
          ++i;
        while (i2 < locations.length && !locations[i2].equals(f2.location))
          ++i2;
        return i < i2;
      }
    }

    /**
     * siehe {@link UNO#findBrowseNodeTreeLeafAndScriptProvider(XBrowseNode,
     * String, String, boolean, String[]))}
     * 
     * @param xScriptProvider
     *          der zuletzt gesehene xScriptProvider
     * @param nameToFind
     *          der zu suchende Name in seine Bestandteile zwischen den Punkten
     *          zerlegt.
     * @param nameToFindLC
     *          wie nameToFind aber alles lowercase.
     * @param prefix
     *          das Prefix in seine Bestandteile zwischen den Punkten zerlegt.
     * @param prefixLC
     *          wie prefix aber alles lowercase.
     * @param found
     *          Liste von {@link FindNode}s mit dem Ergebnis der Suche (anfangs
     *          leere Liste übergeben). Die Sortierung ist so, dass zuerst alle
     *          case-sensitive Matches (also exakte Matches) aufgeführt sind,
     *          sortiert gemäss location und dann alle case-insensitive Matches
     *          sortiert gemäss location. Falls <code>location == null</code>,
     *          so wird nur nach case-sensitive und case-insenstive sortiert,
     *          innerhalb dieser Gruppen jedoch nicht mehr.
     * @return die Anzahl der Rekursionsstufen, die beendet werden sollen. Zum
     *         Beispiel heisst ein Rückgabewert von 1, dass die aufrufende
     *         Funktion ein <code>return 0</code> machen soll.
     * 
     * @author bnk
     */
    public static int findBrowseNodeTreeLeavesAndScriptProviders(
        BrowseNode node, List<String> prefix, List<String> prefixLC,
        String[] nameToFind, String[] nameToFindLC, String[] location,
        XScriptProvider xScriptProvider, List<Utils.FindNode> found)
    {
      String name = node.getName();
      String nameLC = name.toLowerCase();

      XScriptProvider xsc = node.as(XScriptProvider.class);
      if (xsc != null)
        xScriptProvider = xsc;

      Iterator<BrowseNode> iter = node.children();
      if (!iter.hasNext())
      {
        /*
         * Falls der Knoten nicht vom Typ SCRIPT ist, interessiert er uns nicht.
         * Auch wenn wir davon ausgehen können, dass alle Geschwister ebenfalls
         * keine SCRIPTS sind, dürfen wir nicht mehrere Stufen nach oben gehen,
         * da die Geschwister CONTAINER sein können.
         */
        if (node.getType() != BrowseNodeTypes.SCRIPT)
          return 0;

        /*
         * Falls die location des aktuellen Knotens nicht in der erlaubten Liste
         * ist, können wir gleich 2 Ebenen aufsteigen (d.h zur nächsten
         * Library), weil wir davon ausgehen können, dass innerhalb einer
         * Library alle Skripte die selbe Location haben.
         */
        String nodeLocation = getLocation(node);
        if (location != null && !stringInArray(nodeLocation, location))
        {
          return 2;
        }

        /*
         * Wenn das Präfix schon nicht zu nameToFind passt, dann hat es keinen
         * Sinn, alle Skripte des Moduls durchzuiterieren, weil keines davon
         * passen wird. Wir bestimmen, wieviele Rekursionsstufen wir verlassen
         * können. 0 => Präfix passt zur nameToFind 1 => Wir versuchen das
         * nächste Modul in der selben Library, d.h. letzte Präfix-Komponente
         * passt nicht 2 => Wir versuchen die nächste Library, d.h. die letzten
         * 2 Präfix-Komponenten passen nicht. Mehr Ebenen zu verlassen erlauben
         * wir nicht, da es möglich sein kann, dass in verschiedenen Libraries
         * sich die Skripte auf verschiedener Ebene befinden. Im Prinzip ist
         * schon die Annahme, dass sich innerhalb einer Library alle Skripte auf
         * der selben Ebene befinden etwas gewagt. Für Basic ist sie sicher
         * richtig, aber OOo erlaubt noch viele andere Skriptsprachen. Technisch
         * gesehen müsste diese Optimierung die Programmiersprache
         * miteinbeziehen. Im Falle von Basic könnte man vermutlich noch
         * aggressiver sein. Im Falle anderer Sprachen müsste man wohl noch
         * konservativer sein.
         */
        int nMPC = 0;
        if (!prefixLC.isEmpty() && nameToFindLC.length >= 2
            && !prefixLC.get(prefixLC.size() - 1)
                .equals(nameToFindLC[nameToFindLC.length - 2]))
        {
          nMPC = 1;
        }
        if (prefixLC.size() >= 2 && nameToFindLC.length >= 3
            && !prefixLC.get(prefixLC.size() - 2)
                .equals(nameToFindLC[nameToFindLC.length - 3]))
        {
          // ACHTUNG! Hier wird nMPC immer auf 2 gesetzt, nicht inkrementiert.
          // Wenn der Libraryname nicht passt ist es egal, ob der Modulname
          // übereinstimmt!
          nMPC = 2;
        }
        if (nMPC > 0)
          return nMPC;

        // If the name doesn't even match case-insensitive, try the next
        // sibling.
        if (!nameLC.equals(nameToFindLC[nameToFindLC.length - 1]))
          return 0;

        boolean isCaseCorrect = true;
        prefix.add(name); // ACHTUNG! Muss nachher wieder entfernt werden
        for (int i = nameToFind.length - 1, j = prefix.size() - 1; i >= 0
            && j >= 0; --i, --j)
        {
          if (!nameToFind[i].equals(prefix.get(j)))
          {
            isCaseCorrect = false;
            break;
          }
        }
        prefix.remove(prefix.size() - 1); // wieder entfernen vor dem nächsten
                                          // return

        Utils.FindNode findNode = new FindNode(node.unwrap(), xScriptProvider,
            nodeLocation, isCaseCorrect);
        ListIterator<Utils.FindNode> liter = found.listIterator();
        while (liter.hasNext())
        {
          if (findNode.betterMatchThan(liter.next(), location))
          {
            liter.previous();
            break;
          }
        }
        liter.add(findNode);

        /*
         * ACHTUNG: Wir haben einen passenden Knoten gefunden. Nun könnten wir
         * davon ausgehen, dass es im selben Modul keine weiteren Matches gibt
         * und return 1 machen als Optimierung. Bei BASIC Makros ist dies auch
         * korrekt, aber bei Makros in case-sensitiven Sprachen ist es durchaus
         * möglich, dass im selben Modul mehrere Matches (in unterschiedlicher
         * Gross/Kleinschrift) sind. Mehr als return 1 ist auch bei BASIC nicht
         * drin, weil auch BASIC bei Modul und Bibliotheksnamen case-sensitive
         * ist.
         */
        if ("basic".equalsIgnoreCase(getLanguage(node)))
          return 1;
        else
          return 0;
      } else // if iter.hasNext()
      {
        /*
         * ACHTUNG! Diese Änderungen müssen vor return wieder Rückgängig gemacht
         * werden
         */
        prefix.add(name);
        prefixLC.add(name.toLowerCase());

        while (iter.hasNext())
        {
          BrowseNode child = (BrowseNode) iter.next();

          int retL = findBrowseNodeTreeLeavesAndScriptProviders(child, prefix,
              prefixLC, nameToFind, nameToFindLC, location, xScriptProvider,
              found);
          if (retL > 0)
          {
            prefix.remove(prefix.size() - 1);
            prefixLC.remove(prefixLC.size() - 1);
            return retL - 1;
          }
        }

        prefix.remove(prefix.size() - 1);
        prefixLC.remove(prefixLC.size() - 1);

      }
      return 0;
    }

    /**
     * Returns true iff array contains a String that is equals to str
     * 
     * @author bnk
     */
    private static boolean stringInArray(String str, String[] array)
    {
      for (int i = 0; i < array.length; ++i)
        if (str.equals(array[i]))
          return true;
      return false;
    }

    /**
     * Falls <code>node.URL() == null</code> oder die URL keinen "location="
     * Teil enthält, so wird "" geliefert, ansonsten der "location=" Teil ohne
     * das führende "location=".
     * 
     * @author bnk
     */
    private static String getLocation(BrowseNode node)
    { // T
      return getUrlComponent(node, "location");
    }

    private static String getLanguage(BrowseNode node)
    {
      return getUrlComponent(node, "language");
    }

    private static String getUrlComponent(BrowseNode node, String id)
    {
      String url = node.getURL();
      if (url == null)
        return "";
      int idx = url.indexOf("?" + id + "=");
      if (idx < 0)
        idx = url.indexOf("&" + id + "=");
      if (idx < 0)
        return "";
      idx += 10;
      int idx2 = url.indexOf('&', idx);
      if (idx2 < 0)
        idx2 = url.length();
      return url.substring(idx, idx2);
    }

    /**
     * Wenn <code>provider = null</code>, so wird versucht, einen passenden
     * Provider zu finden.
     * 
     * @author bnk
     */
    static Object executeMacroInternal(String macroName, Object[] args,
        XScriptProvider provider, XBrowseNode root, String[] location)
    { // T
      XBrowseNodeAndXScriptProvider o = UNO
          .findBrowseNodeTreeLeafAndScriptProvider(root, "", macroName, false,
              location);

      if (provider == null)
        provider = o.XScriptProvider;
      XScript script;
      try
      {
        String uri = (String) UNO.getProperty(o.XBrowseNode, "URI");
        script = provider.getScript(uri);
      } catch (Exception x)
      {
        throw new RuntimeException(
            "Objekt " + macroName + " nicht gefunden oder ist kein Skript");
      }

      short[][] aOutParamIndex = new short[][] { new short[0] };
      Object[][] aOutParam = new Object[][] { new Object[0] };
      try
      {
        return script.invoke(args, aOutParamIndex, aOutParam);
      } catch (Exception x)
      {
        x.printStackTrace();
        throw new RuntimeException(
            "Fehler bei invoke() von Makro " + macroName);
      }
    }
  }