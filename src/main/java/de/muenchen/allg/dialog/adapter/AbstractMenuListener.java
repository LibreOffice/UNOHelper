package de.muenchen.allg.dialog.adapter;

import com.sun.star.awt.MenuEvent;
import com.sun.star.awt.XMenuListener;
import com.sun.star.lang.EventObject;

/**
 * Provides default implementations of standard methods for the {@link XMenuListener}.
 */
public abstract class AbstractMenuListener implements XMenuListener
{

  @Override
  public void disposing(EventObject event)
  {
    // default implementation
  }

  @Override
  public void itemActivated(MenuEvent event)
  {
    // default implementation
  }

  @Override
  public void itemDeactivated(MenuEvent event)
  {
    // default implementation
  }

  @Override
  public void itemHighlighted(MenuEvent event)
  {
    // default implementation
  }

  @Override
  public void itemSelected(MenuEvent event)
  {
    // default implementation
  }
}
