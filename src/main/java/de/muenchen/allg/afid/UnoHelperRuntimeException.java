package de.muenchen.allg.afid;

/**
 * Runtime exception in UnoHelper.
 */
public class UnoHelperRuntimeException extends RuntimeException
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
  public UnoHelperRuntimeException(String msg, Throwable e)
  {
    super(msg, e);
  }

  /**
   * Create new exception.
   *
   * @param msg
   *          The message of the exception.
   */
  public UnoHelperRuntimeException(String msg)
  {
    super(msg);
  }

  /**
   * Create new exception.
   *
   * @param e
   *          The cause of the exception.
   */
  public UnoHelperRuntimeException(Throwable e)
  {
    super(e);
  }
}
