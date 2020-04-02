package de.muenchen.allg.dialog.adapter;

import com.sun.star.awt.WindowEvent;
import com.sun.star.awt.XWindowListener;
import com.sun.star.lang.EventObject;

/**
 * Provides default implementations of standard methods for the {@link XWindowListener}.
 */
public abstract class AbstractWindowListener implements XWindowListener
{

  @Override
  public void disposing(EventObject event)
  {
    // default implementation
  }

  @Override
  public void windowHidden(EventObject event)
  {
    // default implementation
  }

  @Override
  public void windowMoved(WindowEvent event)
  {
    // default implementation
  }

  @Override
  public void windowResized(WindowEvent event)
  {
    // default implementation
  }

  @Override
  public void windowShown(EventObject event)
  {
    // default implementation
  }

}
