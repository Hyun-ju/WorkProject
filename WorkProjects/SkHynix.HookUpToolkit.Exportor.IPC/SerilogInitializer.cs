using Serilog;

namespace SkHynix.HookUpToolkit.Exportor.IPC {
  public class SerilogInitializer {
    /// <summary>
    /// 지정 경로에 로깅 하도록 초기화
    /// </summary>
    /// <param name="logDirectory">로그 디렉토리 경로</param>
    public static void Initialize(string logDirectory) {
      if (System.IO.Directory.Exists(logDirectory) == false) {
        System.IO.Directory.CreateDirectory(logDirectory);
      }

      Log.Logger = new LoggerConfiguration()
          .MinimumLevel.Verbose()
          .WriteTo.Logger(delegate (LoggerConfiguration lc) {
            lc.MinimumLevel.Warning().WriteTo.Debug()

#if DEBUG
                .WriteTo.File(System.IO.Path.Combine(logDirectory, $"log.txt"),
                    rollingInterval: RollingInterval.Hour);
#else
                .WriteTo.File(System.IO.Path.Combine(logDirectory, "log.txt"),
                    rollingInterval: RollingInterval.Day);
#endif
          })
         .WriteTo.Logger(delegate (LoggerConfiguration lc) {
           lc.MinimumLevel.Information()
              .Filter.ByIncludingOnly(delegate (Serilog.Events.LogEvent evt) {
                return (evt.Level <= Serilog.Events.LogEventLevel.Information);
              }).WriteTo.Debug().WriteTo.File(
                  System.IO.Path.Combine(logDirectory, $"Information.txt"),
                  rollingInterval: RollingInterval.Hour);
         })
#if DEBUG
         .WriteTo.Logger(delegate (LoggerConfiguration lc) {
           lc.MinimumLevel.Debug()
              .Filter.ByIncludingOnly(delegate (Serilog.Events.LogEvent evt) {
                return (evt.Level <= Serilog.Events.LogEventLevel.Debug);
              }).WriteTo.Debug().WriteTo.File(
                  System.IO.Path.Combine(logDirectory, $"Debug.txt"),
                  rollingInterval: RollingInterval.Hour);
         })
          .WriteTo.Logger(delegate (LoggerConfiguration lc) {
            lc.MinimumLevel.Verbose()
                .Filter.ByIncludingOnly(delegate (Serilog.Events.LogEvent evt) {
                  return (evt.Level <= Serilog.Events.LogEventLevel.Verbose);
                }).WriteTo.Debug().WriteTo.File(
                    System.IO.Path.Combine(logDirectory, $"Verbose.txt"),
                    rollingInterval: RollingInterval.Hour);
          })
#endif
                .CreateLogger();

#if DEBUG
      Log.Information("Initialized Serilog!!");
#endif
    }

    /// <summary>
    /// Dll 위치의 log 폴더로 경로 생성
    /// </summary>
    public static void Initialize() {
      var path = System.IO.Path.GetDirectoryName(
          typeof(SerilogInitializer).Assembly.Location);
      Initialize(System.IO.Path.Combine(path, "log"));
    }
  }
}
