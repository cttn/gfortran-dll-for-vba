!!Compiled with:
!!  i686-w64-mingw32-gfortran -shared exVBA.f90 -fno-underscoring -o exVBA.dll

!!Declared in VBA as:
!!  Declare PtrSafe Function exscalar Lib "D:\exVBA.dll" Alias "exscalar@4" (x As Double) As Double 
function exscalar(x)  !fortra ignores capitalization, but VBA doesn't. So I need a capitalized alias.
  implicit none
!GCC$ ATTRIBUTES STDCALL, DLLEXPORT :: exscalar
  real(kind(0d0)) :: exscalar
  real(kind(0d0)),intent(in) :: x
  exscalar = 1.05d0*x + 0.11d0
end function exscalar

!!Declared in VBA as:
!!  Declare PtrSafe Function exarrays Lib "D:\exVBA.dll" Alias "exarrays@8" (ByVal n as Long, x As Double) As Double 
function exarrays(n,x)  !fortra ignores capitalization, but VBA doesn't. So I need a capitalized alias.
  implicit none
  !GCC$ ATTRIBUTES STDCALL, DLLEXPORT :: exarrays
  real(kind(0d0)) :: exarrays
  integer, intent(in), value :: n  !Up bound passed by value
  real(kind(0d0)),intent(in) :: x(n)  !Array passed by reference
  exarrays = x(1) *0.05d0 - 1.d5
end function exarrays

!!Declared in VBA as:
!!  Declare Sub subarrays Lib "D:\exVBA.dll" Alias "subarrays@12" (ByVal n as Long, x As Double, res As Double)
subroutine subarrays(n,x,res)  !fortra ignores capitalization, but VBA doesn't. So I need a capitalized alias.
  implicit none
  !GCC$ ATTRIBUTES STDCALL, DLLEXPORT :: subarrays
  integer, intent(in), value :: n  !Up bound passed by value
  real(kind(0d0)),intent(in) :: x(n)  !Array passed by reference
  real(kind(0d0)),intent(out) :: res(n)  !Array passed by reference
  res(:) = x(:) - 1.d1
end subroutine subarrays