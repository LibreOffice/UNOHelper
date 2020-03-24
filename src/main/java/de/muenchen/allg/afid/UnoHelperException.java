package de.muenchen.allg.afid;

public class UnoHelperException extends Exception
{
  private static final long serialVersionUID = 1L;

  public UnoHelperException(String msg, Throwable e)
  {
    super(msg, e);
  }

  public UnoHelperException(String msg)
  {
    super(msg);
  }

  public UnoHelperException(Throwable e)
  {
    super(e);
  }
}
