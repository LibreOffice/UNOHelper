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
