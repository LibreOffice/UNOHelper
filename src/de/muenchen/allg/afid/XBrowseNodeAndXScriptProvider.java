package de.muenchen.allg.afid;

import com.sun.star.script.browse.XBrowseNode;
import com.sun.star.script.provider.XScriptProvider;

public class XBrowseNodeAndXScriptProvider
  {
    private XBrowseNode xBrowseNode = null;
    private XScriptProvider xScriptProvider = null;

    public XBrowseNodeAndXScriptProvider()
    {
    }

    public XBrowseNodeAndXScriptProvider(XBrowseNode xBrowseNode,
        XScriptProvider xScriptProvider)
    {
      this.xBrowseNode = xBrowseNode;
      this.xScriptProvider = xScriptProvider;
    }

    public XBrowseNode getXBrowseNode()
    {
      return xBrowseNode;
    }

    public XScriptProvider getXScriptProvider()
    {
      return xScriptProvider;
    }
  }