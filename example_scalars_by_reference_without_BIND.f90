!! Trying without BIND C
function fooSimple(x)  !fortra ignores capitalization, but VBA doesn't. So I need a capitalized alias.
  implicit none
  !GCC$ ATTRIBUTES STDCALL, DLLEXPORT, ALIAS: 'FOOSIMPLE' :: fooSimple
  real(kind(0d0)) :: fooSimple
  real(kind(0d0)),intent(in) :: x

  fooSimple = 1.05d0*x

end function fooSimple

