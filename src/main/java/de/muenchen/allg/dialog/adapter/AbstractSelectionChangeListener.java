package de.muenchen.allg.dialog.adapter;

import com.sun.star.lang.EventObject;
import com.sun.star.view.XSelectionChangeListener;

/**
 * Provides default implementations of standard methods for the {@link XSelectionChangeListener}.
 */
@FunctionalInterface
public interface AbstractSelectionChangeListener extends XSelectionChangeListener
{

  @Override
  public default void disposing(EventObject arg0)
  {
    // default implementation
  }

}
