package de.muenchen.allg.dialog.adapter;

import com.sun.star.awt.MouseEvent;
import com.sun.star.awt.XMouseListener;
import com.sun.star.lang.EventObject;

/**
 * Provides default implementations of standard methods for the {@link XMouseListener}.
 */
public abstract class AbstractMouseListener implements XMouseListener
{

  @Override
  public void disposing(EventObject event)
  {
    // default implementation
  }

  @Override
  public void mouseEntered(MouseEvent event)
  {
    // default implementation
  }

  @Override
  public void mouseExited(MouseEvent event)
  {
    // default implementation
  }

  @Override
  public void mousePressed(MouseEvent event)
  {
    // default implementation
  }

  @Override
  public void mouseReleased(MouseEvent event)
  {
    // default implementation
  }

}
