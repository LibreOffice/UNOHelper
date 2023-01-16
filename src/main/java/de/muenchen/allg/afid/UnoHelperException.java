/*-
 * #%L
 * UNOHelper
 * %%
 * Copyright (C) 2005 - 2023 Landeshauptstadt München
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
package de.muenchen.allg.afid;

/**
 * Exception occurred in UnoHelper.
 */
public class UnoHelperException extends Exception
{
  private static final long serialVersionUID = 1L;

  /**
   * Create new exception.
   *
   * @param msg
   *          The message of the exception.
   * @param e
   *          The cause.
   */
  public UnoHelperException(String msg, Throwable e)
  {
    super(msg, e);
  }

  /**
   * Create new exception.
   *
   * @param msg
   *          The message of the exception.
   */
  public UnoHelperException(String msg)
  {
    super(msg);
  }

  /**
   * Create new exception.
   *
   * @param e
   *          The cause of the exception.
   */
  public UnoHelperException(Throwable e)
  {
    super(e);
  }
}
