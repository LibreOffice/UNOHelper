package de.muenchen.allg.dialog.adapter;

import com.sun.star.lang.EventObject;
import com.sun.star.view.XSelectionChangeListener;

/**
 * Provides default implementations of standard methods for the {@link XSelectionChangeListener}.
 */
public abstract class AbstractSelectionChangeListener implements XSelectionChangeListener
{

  @Override
  public void disposing(EventObject arg0)
  {
    // default implementation
  }

  @Override
  public void selectionChanged(EventObject arg0)
  {
    // default implementation
  }

}
