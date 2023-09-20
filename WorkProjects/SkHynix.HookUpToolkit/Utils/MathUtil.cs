using System;

namespace SkHynix.HookUpToolkit.Utils {
  public static class MathUtil {
    public const double Tolerance = 1e-5;

    /// <summary>
    /// 공차를 사용한 값 비교
    /// </summary>
    /// <param name="tolerance">기본 값 1e-5</param>
    /// <returns></returns>
    public static bool IsAlmostEqualTo(this double a, double b, double tolerance = Tolerance) {
      return Math.Abs(a - b) < tolerance;
    }

    /// <summary>
    /// 공차를 사용한 값 비교
    /// </summary>
    /// <param name="tolerance">기본 값 1e-5</param>
    /// <returns></returns>
    public static bool IsAlmostEqualTo(this float a, double b, double tolerance = Tolerance) {
      return Math.Abs(a - b) < tolerance;
    }
  }
}
