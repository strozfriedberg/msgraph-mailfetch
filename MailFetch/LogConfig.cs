/*
 * Copyright 2025 LevelBlue
 *
 * Licensed under the Apache License, Version 2.0 (the "License");
 * you may not use this file except in compliance with the License.
 * You may obtain a copy of the License at
 *
 *     http://www.apache.org/licenses/LICENSE-2.0
 *
 * Unless required by applicable law or agreed to in writing, software
 * distributed under the License is distributed on an "AS IS" BASIS,
 * WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 * See the License for the specific language governing permissions and
 * limitations under the License.
 */

using System;
using System.IO;

namespace MailFetch
{
    public sealed class LogConfig : IDisposable
    {
        public static readonly string[] LogContexts = new string[] { "Collection", "UserCollection", "Result" };
        private readonly NLog.Config.LoggingConfiguration _loggingConfig = new NLog.Config.LoggingConfiguration();
        private readonly string? _rootLogFilename;

        private static string GetExceptionLayout(string exceptionFormat)
        {
            return $"${{exception:format={exceptionFormat}:exceptionDataSeparator=\r\n:innerFormat={exceptionFormat}:innerExceptionSeparator=\r\n:maxInnerExceptionLevel=99999}}";
        }

        private static string GetLoggerLayout(string exceptionFormat)
        {
            var exceptionLayout = GetExceptionLayout(exceptionFormat);
            return "[${longdate:universalTime=True}Z] "
                   + "${level:upperCase=true:other:padding=-5} "
                   + "${callsite:className=False:includeNamespace=False:fileName=True:includeSourcePath=False:methodName=False:skipFrames=0:other:padding=-20} "
                   + "${when:when='${mdlc:item=UserCollection}'=='':inner=${mdlc:item=Collection}:else=${mdlc:item=UserCollection}}"
                   + "${when:when='${mdlc:item=Result}'=='':else=/${mdlc:item=Result}} "
                   + "${message} "
                   + exceptionLayout;
        }

        public LogConfig(CLIOptions options)
        {
            _rootLogFilename = $"mail-fetch_{options.Started:yyyyMMddHHmmssfffffff}.log";

            Directory.CreateDirectory(options.Output);
            AddLogger("globalLogger", Path.Combine(options.Output, "logs", _rootLogFilename));

            // Log to console
            var logConsole = new NLog.Targets.ConsoleTarget("consoleLogger") { Layout = GetLoggerLayout("Type,Message,Data") };
            _loggingConfig.AddRule(NLog.LogLevel.Info, NLog.LogLevel.Fatal, logConsole);

            // Apply config
            NLog.LogManager.Configuration = _loggingConfig;
        }

        private void AddLogger(string name, string filename)
        {
            Directory.CreateDirectory(Path.GetDirectoryName(filename)!);
            var logFile = new NLog.Targets.FileTarget(name) { FileName = filename, Layout = GetLoggerLayout("Type,Message,Data,StackTrace") };
            _loggingConfig.AddRule(NLog.LogLevel.Trace, NLog.LogLevel.Fatal, logFile);
            NLog.LogManager.ReconfigExistingLoggers();
        }

        public void Dispose()
        {
            NLog.LogManager.Shutdown();
        }
    }
}
