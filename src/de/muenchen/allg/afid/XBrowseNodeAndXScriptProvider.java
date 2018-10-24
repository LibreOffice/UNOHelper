package de.muenchen.allg.afid;

import com.sun.star.script.browse.XBrowseNode;
import com.sun.star.script.provider.XScriptProvider;

public class XBrowseNodeAndXScriptProvider
  {
    public XBrowseNode XBrowseNode = null;
    public XScriptProvider XScriptProvider = null;

    public XBrowseNodeAndXScriptProvider()
    {
    }

    public XBrowseNodeAndXScriptProvider(XBrowseNode xBrowseNode,
        XScriptProvider xScriptProvider)
    {
      this.XBrowseNode = xBrowseNode;
      this.XScriptProvider = xScriptProvider;
    }
  }