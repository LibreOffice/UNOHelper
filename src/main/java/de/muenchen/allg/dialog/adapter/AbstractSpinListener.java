package de.muenchen.allg.dialog.adapter;

import com.sun.star.awt.SpinEvent;
import com.sun.star.awt.XSpinListener;
import com.sun.star.lang.EventObject;

/**
 * Provides default implementations of standard methods for the {@link XSpinListener}.
 */
public abstract class AbstractSpinListener implements XSpinListener
{

  @Override
  public void disposing(EventObject event)
  {
    // default implementation
  }

  @Override
  public void down(SpinEvent event)
  {
    // default implementation
  }

  @Override
  public void first(SpinEvent event)
  {
    // default implementation
  }

  @Override
  public void last(SpinEvent event)
  {
    // default implementation
  }

  @Override
  public void up(SpinEvent event)
  {
    // default implementation
  }

}
