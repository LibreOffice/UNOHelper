package de.muenchen.allg.dialog.adapter;

import com.sun.star.awt.XActionListener;
import com.sun.star.lang.EventObject;

/**
 * Provides default implementations of standard methods for the {@link XActionListener}.
 */
@FunctionalInterface
public interface AbstractActionListener extends XActionListener
{

  @Override
  public default void disposing(EventObject event)
  {
    // default implementation
  }
}
