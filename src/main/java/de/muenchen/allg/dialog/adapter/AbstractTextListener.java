package de.muenchen.allg.dialog.adapter;

import com.sun.star.awt.XTextListener;
import com.sun.star.lang.EventObject;

/**
 * Provides default implementations of standard methods for the {@link XTextListener}.
 */
@FunctionalInterface
public interface AbstractTextListener extends XTextListener
{
  @Override
  default void disposing(EventObject event)
  {
    // default implementation
  }
}
