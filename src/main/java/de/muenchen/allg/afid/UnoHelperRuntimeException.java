package de.muenchen.allg.afid;

public class UnoHelperRuntimeException extends RuntimeException
{
  private static final long serialVersionUID = 1L;

  public UnoHelperRuntimeException(String msg, Throwable e)
  {
    super(msg, e);
  }

  public UnoHelperRuntimeException(String msg)
  {
    super(msg);
  }

  public UnoHelperRuntimeException(Throwable e)
  {
    super(e);
  }
}
