function fooSimple(x) Bind(c) !fortra ignores capitalization, but VBA doesn't. So I need a capitalized alias.
  use ISO_BIND_C
  implicit none
  !GCC$ ATTRIBUTES STDCALL, DLLEXPORT, ALIAS: 'FOOSIMPLE' :: fooSimple
  real(c_double) :: fooSimple
  real(c_double), value, intent(in) :: x !! By Value!
  fooSimple = 1.05d0*x
end function fooSimple

