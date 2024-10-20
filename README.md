public bool IsCVErr(object obj)
{
    if (obj is int)
    {
        switch ((int)obj)
        {
            case (int)CVErrEnum.ErrDiv0:
            case (int)CVErrEnum.ErrNA:
            case (int)CVErrEnum.ErrName:
            case (int)CVErrEnum.ErrNull:
            case (int)CVErrEnum.ErrNum:
            case (int)CVErrEnum.ErrRef:
            case (int)CVErrEnum.ErrValue:
                return true;
            default:
                return false;
        }
    }
    return false;
}
