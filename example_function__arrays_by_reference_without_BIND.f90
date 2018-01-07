function exArrays(n,x)  !fortra ignores capitalization, but VBA doesn't. So I need a capitalized alias.
  implicit none
  !GCC$ ATTRIBUTES STDCALL, DLLEXPORT, ALIAS: 'EXARRAYS' :: exArrays
  real(kind(0d0)) :: exArrays
  integer, intent(in), value :: n  !Up bound passed by value
  real(kind(0d0)),intent(in) :: x(n)  !Array passed by reference

  exArrays = x(1) * x(n)
end function exArrays

