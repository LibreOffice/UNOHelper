package de.muenchen.allg.dialog.adapter;

import com.sun.star.awt.XItemListener;
import com.sun.star.lang.EventObject;

/**
 * Provides default implementations of standard methods for the {@link XItemListener}.
 */
@FunctionalInterface
public interface AbstractItemListener extends XItemListener
{
  @Override
  default void disposing(EventObject event)
  {
    // default implementation
  }
}
