/*-
 * #%L
 * WollMux
 * %%
 * Copyright (C) 2005 - 2022 Landeshauptstadt München
 * %%
 * Licensed under the EUPL, Version 1.1 or – as soon they will be
 * approved by the European Commission - subsequent versions of the
 * EUPL (the "Licence");
 *
 * You may not use this work except in compliance with the Licence.
 * You may obtain a copy of the Licence at:
 *
 * http://ec.europa.eu/idabc/eupl5
 *
 * Unless required by applicable law or agreed to in writing, software
 * distributed under the Licence is distributed on an "AS IS" basis,
 * WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 * See the Licence for the specific language governing permissions and
 * limitations under the Licence.
 * #L%
 */
package de.muenchen.allg.ui.layout;

import org.apache.commons.lang3.tuple.Pair;

import com.sun.star.awt.PosSize;
import com.sun.star.awt.Rectangle;
import com.sun.star.awt.Size;
import com.sun.star.awt.XControl;
import com.sun.star.awt.XLayoutConstrains;
import com.sun.star.awt.XWindow;
import com.sun.star.awt.XWindow2;

import de.muenchen.allg.afid.UNO;

/**
 * A layout with one control.
 */
public class ControlLayout implements Layout
{

  /**
   * The control as {@link XWindow}.
   */
  private XWindow control;

  /**
   * The control as {@link XWindow2}.
   */
  private XWindow2 control2;

  /**
   * The control as {@link XLayoutContrains}.
   */
  private XLayoutConstrains constrains;

  /**
   * The initial height of the control.
   */
  private int height;

  /**
   * The maximal height of the control.
   */
  private int maxHeight = Integer.MAX_VALUE;

  /**
   * A new layout for a single control.
   *
   * @param control
   *          The control as {@link XWindow}.
   */
  public ControlLayout(XWindow control)
  {
    this.control = control;
    this.control2 = UNO.XWindow2(control);
    this.constrains = UNO.XLayoutConstrains(control);
    height = control.getPosSize().Height;
  }

  /**
   * A new layout for a single control with a maximal height.
   *
   * @param control
   *          The control as {@link XWindow}
   * @param maxHeight
   *          The maximal height.
   */
  public ControlLayout(XWindow control, int maxHeight)
  {
    this(control);
    this.maxHeight = maxHeight;
  }

  /**
   * A new layout for a single control.
   *
   * @param control
   *          The control as {@link XControl}.
   */
  public ControlLayout(XControl control)
  {
    this(UNO.XWindow(control));
  }

  @Override
  public Pair<Integer, Integer> layout(Rectangle rect)
  {
    int h = height;
    int w = rect.Width;
    if (constrains != null)
    {
      Size size = constrains.calcAdjustedSize(new Size(rect.Width, height));
      h = Integer.max(h, size.Height);
    }
    h = Integer.min(h, maxHeight);
    control.setPosSize(rect.X, rect.Y, w, h, PosSize.POSSIZE);
    if (UNO.XWindow2(control) != null && !UNO.XWindow2(control).isVisible())
    {
      return Pair.of(0, 0);
    }
    return Pair.of(h, w);
  }

  /**
   * Unsupported operation.
   */
  @Override
  public void addLayout(Layout layout, int space)
  {
    throw new UnsupportedOperationException();
  }

  @Override
  public int getHeightForWidth(int width)
  {
    int h = 0;
    if (control2 != null && control2.isVisible())
    {
      h = height;
      if (constrains != null)
      {
        Size size = constrains.calcAdjustedSize(new Size(width, height));
        h = Integer.min(Integer.max(h, size.Height), maxHeight);
      }
    }
    return h;
  }

  @Override
  public int getMinimalWidth(int maxWidth)
  {
    int w = 0;
    if (constrains != null)
    {
      Size size = constrains.getMinimumSize();
      w = size.Width;
    }
    return Integer.min(maxWidth, w);
  }

  @Override
  public boolean isVisible()
  {
    if (control2 != null)
    {
      return control2.isVisible();
    }
    return true;
  }

  @Override
  public int size()
  {
    return 1;
  }
}
