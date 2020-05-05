package de.muenchen.allg.dialog.adapter;

import com.sun.star.awt.tab.XTabPageContainerListener;
import com.sun.star.lang.EventObject;

/**
 * Provides default implementations of standard methods for the {@link XTabPageContainerListener}.
 */
@FunctionalInterface
public interface AbstractTabPageContainerListener extends XTabPageContainerListener
{

  @Override
  public default void disposing(EventObject event)
  {
    // default implementation
  }
}
