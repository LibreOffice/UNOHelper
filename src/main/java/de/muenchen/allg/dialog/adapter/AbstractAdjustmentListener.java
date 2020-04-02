package de.muenchen.allg.dialog.adapter;

import com.sun.star.awt.XAdjustmentListener;
import com.sun.star.lang.EventObject;

/**
 * Provides default implementations of standard methods for the {@link XAdjustmentListener}.
 */
public interface AbstractAdjustmentListener extends XAdjustmentListener
{

  @Override
  public default void disposing(EventObject event)
  {
    // default implementation
  }
}
