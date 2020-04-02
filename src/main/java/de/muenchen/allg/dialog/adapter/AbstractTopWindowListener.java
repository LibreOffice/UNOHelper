package de.muenchen.allg.dialog.adapter;

import com.sun.star.awt.XTopWindowListener;
import com.sun.star.lang.EventObject;

/**
 * Provides default implementations of standard methods for the {@link XTopWindowListener}.
 */
public abstract class AbstractTopWindowListener implements XTopWindowListener
{

  @Override
  public void disposing(EventObject arg0)
  {
    // default implementation
  }

  @Override
  public void windowActivated(EventObject arg0)
  {
    // default implementation
  }

  @Override
  public void windowClosed(EventObject arg0)
  {
    // default implementation
  }

  @Override
  public void windowClosing(EventObject arg0)
  {
    // default implementation
  }

  @Override
  public void windowDeactivated(EventObject arg0)
  {
    // default implementation
  }

  @Override
  public void windowMinimized(EventObject arg0)
  {
    // default implementation
  }

  @Override
  public void windowNormalized(EventObject arg0)
  {
    // default implementation
  }

  @Override
  public void windowOpened(EventObject arg0)
  {
    // default implementation
  }

}
