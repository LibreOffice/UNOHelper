package de.muenchen.allg.dialog.adapter;

import com.sun.star.lang.EventObject;
import com.sun.star.util.CloseVetoException;
import com.sun.star.util.XCloseListener;

/**
 * Provides default implementations of standard methods for the {@link XCloseListener}.
 */
public abstract class AbstractCloseListener implements XCloseListener
{

  @Override
  public void disposing(EventObject arg0)
  {
    // default implementation
  }

  @Override
  public void notifyClosing(EventObject arg0)
  {
    // default implementation
  }

  @Override
  public void queryClosing(EventObject arg0, boolean arg1) throws CloseVetoException
  {
    // default implementation
  }

}
