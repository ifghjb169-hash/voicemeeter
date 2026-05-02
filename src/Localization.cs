using System;
using System.Collections.Generic;

namespace SoundDeviceSwitcher
{
    internal sealed class LanguageOption
    {
        public LanguageOption(string code, string displayName)
        {
            Code = code;
            DisplayName = displayName;
        }

        public string Code { get; private set; }
        public string DisplayName { get; private set; }

        public override string ToString()
        {
            return DisplayName;
        }
    }

    internal static class Localizer
    {
        public const string DefaultLanguage = "zh";

        public static readonly LanguageOption[] LanguageOptions =
        {
            new LanguageOption("zh", "中文"),
            new LanguageOption("en", "English"),
            new LanguageOption("pt", "Português"),
            new LanguageOption("es", "Español")
        };

        private static readonly Dictionary<string, Dictionary<string, string>> Strings =
            new Dictionary<string, Dictionary<string, string>>(StringComparer.OrdinalIgnoreCase)
            {
                {
                    "zh", new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
                    {
                        {"tab.system", "系统默认设备"},
                        {"tab.voicemeeter", "Voicemeeter"},
                        {"tab.language", "语言"},
                        {"button.refresh", "刷新设备"},
                        {"button.vmDefaults", "应用 Voicemeeter 默认路由"},
                        {"button.restart", "重启音频引擎"},
                        {"radio.physicalOnly", "只显示物理/耳机/内置设备"},
                        {"radio.showAll", "显示全部设备"},
                        {"group.output", "播放 / 输出"},
                        {"group.input", "录音 / 输入"},
                        {"group.volumeInput", "输入设备音量"},
                        {"group.volumeOutput", "输出设备音量"},
                        {"group.vmInput", "Hardware Input 1（麦克风）"},
                        {"group.vmOutput", "A1 Hardware Out（扬声器/耳机）"},
                        {"column.output", "输出设备"},
                        {"column.input", "输入设备"},
                        {"column.driverInterface", "驱动/接口"},
                        {"column.status", "状态"},
                        {"column.driver", "驱动"},
                        {"column.hardwareId", "硬件 ID"},
                        {"language.title", "界面语言"},
                        {"language.description", "选择后会立即更新界面，并保存到下次启动。"},
                        {"language.saved", "界面语言已切换为：{0}"},
                        {"volume.none", "没有可调节音量的物理设备"},
                        {"volume.unavailable", "不可用"},
                        {"status.systemRead", "已读取 {0} 个输出设备、{1} 个输入设备。"},
                        {"status.systemReadFailed", "读取系统声音设备失败：{0}"},
                        {"status.startupDefaultsApplied", "启动时已应用 Voicemeeter 默认路由：{0}"},
                        {"status.startupDefaultsMissing", "启动时没有找到完整的 Voicemeeter 默认设备：{0}"},
                        {"status.vmDefaultOutput", "默认输出 -> {0}"},
                        {"status.vmDefaultInput", "默认输入 -> {0}"},
                        {"status.vmDefaultOutputMissing", "没有找到 VoiceMeeter Input 默认输出设备"},
                        {"status.vmDefaultInputMissing", "没有找到 VoiceMeeter Output / B1 默认输入设备"},
                        {"status.vmDefaultsFailed", "应用 Voicemeeter 默认路由失败：{0}"},
                        {"status.defaultOutputSet", "默认输出已设置为：{0}"},
                        {"status.defaultInputSet", "默认输入已设置为：{0}"},
                        {"status.volumeSet", "{0} 音量已设置为 {1}%"},
                        {"status.volumeSetFailed", "设置 {0} 音量失败：{1}"},
                        {"status.systemSetFailed", "设置系统默认设备失败：{0}"},
                        {"message.systemSetFailed", "设置系统默认设备失败"},
                        {"status.vmConnected", "{0} 已连接；{1}；DLL: {2}"},
                        {"status.modeAll", "全部设备"},
                        {"status.modePhysical", "物理设备视图"},
                        {"status.vmReadFailed", "Voicemeeter 连接或读取失败：{0}"},
                        {"status.vmInputSet", "Hardware Input 1 已设置为：{0} - {1}"},
                        {"status.vmOutputSet", "A1 已设置为：{0} - {1}"},
                        {"status.vmSetFailed", "设置 Voicemeeter 设备失败：{0}"},
                        {"message.vmSetFailed", "设置 Voicemeeter 设备失败"},
                        {"status.vmRestarted", "已向 Voicemeeter 发送重启音频引擎命令。"},
                        {"status.vmRestartFailed", "重启音频引擎失败：{0}"}
                    }
                },
                {
                    "en", new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
                    {
                        {"tab.system", "System Default"},
                        {"tab.voicemeeter", "Voicemeeter"},
                        {"tab.language", "Language"},
                        {"button.refresh", "Refresh"},
                        {"button.vmDefaults", "Apply Voicemeeter Defaults"},
                        {"button.restart", "Restart Audio Engine"},
                        {"radio.physicalOnly", "Physical/headset/built-in only"},
                        {"radio.showAll", "Show all devices"},
                        {"group.output", "Playback / Output"},
                        {"group.input", "Recording / Input"},
                        {"group.volumeInput", "Input Device Volume"},
                        {"group.volumeOutput", "Output Device Volume"},
                        {"group.vmInput", "Hardware Input 1 (Microphone)"},
                        {"group.vmOutput", "A1 Hardware Out (Speakers/Headset)"},
                        {"column.output", "Output Device"},
                        {"column.input", "Input Device"},
                        {"column.driverInterface", "Driver/Interface"},
                        {"column.status", "Status"},
                        {"column.driver", "Driver"},
                        {"column.hardwareId", "Hardware ID"},
                        {"language.title", "Interface language"},
                        {"language.description", "The interface updates immediately and the choice is saved for next launch."},
                        {"language.saved", "Interface language changed to: {0}"},
                        {"volume.none", "No physical device with adjustable volume"},
                        {"volume.unavailable", "Unavailable"},
                        {"status.systemRead", "Loaded {0} output devices and {1} input devices."},
                        {"status.systemReadFailed", "Failed to read system audio devices: {0}"},
                        {"status.startupDefaultsApplied", "Applied Voicemeeter default route on startup: {0}"},
                        {"status.startupDefaultsMissing", "Could not find the full Voicemeeter default devices on startup: {0}"},
                        {"status.vmDefaultOutput", "Default output -> {0}"},
                        {"status.vmDefaultInput", "Default input -> {0}"},
                        {"status.vmDefaultOutputMissing", "VoiceMeeter Input default output device was not found"},
                        {"status.vmDefaultInputMissing", "VoiceMeeter Output / B1 default input device was not found"},
                        {"status.vmDefaultsFailed", "Failed to apply Voicemeeter default route: {0}"},
                        {"status.defaultOutputSet", "Default output set to: {0}"},
                        {"status.defaultInputSet", "Default input set to: {0}"},
                        {"status.volumeSet", "{0} volume set to {1}%"},
                        {"status.volumeSetFailed", "Failed to set {0} volume: {1}"},
                        {"status.systemSetFailed", "Failed to set system default device: {0}"},
                        {"message.systemSetFailed", "Failed to set system default device"},
                        {"status.vmConnected", "{0} connected; {1}; DLL: {2}"},
                        {"status.modeAll", "all devices"},
                        {"status.modePhysical", "physical device view"},
                        {"status.vmReadFailed", "Failed to connect to or read Voicemeeter: {0}"},
                        {"status.vmInputSet", "Hardware Input 1 set to: {0} - {1}"},
                        {"status.vmOutputSet", "A1 set to: {0} - {1}"},
                        {"status.vmSetFailed", "Failed to set Voicemeeter device: {0}"},
                        {"message.vmSetFailed", "Failed to set Voicemeeter device"},
                        {"status.vmRestarted", "Restart audio engine command sent to Voicemeeter."},
                        {"status.vmRestartFailed", "Failed to restart audio engine: {0}"}
                    }
                },
                {
                    "pt", new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
                    {
                        {"tab.system", "Padrão do Sistema"},
                        {"tab.voicemeeter", "Voicemeeter"},
                        {"tab.language", "Idioma"},
                        {"button.refresh", "Atualizar"},
                        {"button.vmDefaults", "Aplicar padrões do Voicemeeter"},
                        {"button.restart", "Reiniciar motor de áudio"},
                        {"radio.physicalOnly", "Somente físicos/fone/integrados"},
                        {"radio.showAll", "Mostrar todos os dispositivos"},
                        {"group.output", "Reprodução / Saída"},
                        {"group.input", "Gravação / Entrada"},
                        {"group.volumeInput", "Volume dos Dispositivos de Entrada"},
                        {"group.volumeOutput", "Volume dos Dispositivos de Saída"},
                        {"group.vmInput", "Hardware Input 1 (Microfone)"},
                        {"group.vmOutput", "A1 Hardware Out (Alto-falantes/Fone)"},
                        {"column.output", "Dispositivo de Saída"},
                        {"column.input", "Dispositivo de Entrada"},
                        {"column.driverInterface", "Driver/Interface"},
                        {"column.status", "Status"},
                        {"column.driver", "Driver"},
                        {"column.hardwareId", "ID de hardware"},
                        {"language.title", "Idioma da interface"},
                        {"language.description", "A interface muda imediatamente e a escolha fica salva para a próxima abertura."},
                        {"language.saved", "Idioma da interface alterado para: {0}"},
                        {"volume.none", "Nenhum dispositivo físico com volume ajustável"},
                        {"volume.unavailable", "Indisponível"},
                        {"status.systemRead", "{0} dispositivos de saída e {1} dispositivos de entrada carregados."},
                        {"status.systemReadFailed", "Falha ao ler dispositivos de áudio do sistema: {0}"},
                        {"status.startupDefaultsApplied", "Rota padrão do Voicemeeter aplicada ao iniciar: {0}"},
                        {"status.startupDefaultsMissing", "Não foi possível encontrar todos os dispositivos padrão do Voicemeeter ao iniciar: {0}"},
                        {"status.vmDefaultOutput", "Saída padrão -> {0}"},
                        {"status.vmDefaultInput", "Entrada padrão -> {0}"},
                        {"status.vmDefaultOutputMissing", "Dispositivo de saída padrão VoiceMeeter Input não encontrado"},
                        {"status.vmDefaultInputMissing", "Dispositivo de entrada padrão VoiceMeeter Output / B1 não encontrado"},
                        {"status.vmDefaultsFailed", "Falha ao aplicar a rota padrão do Voicemeeter: {0}"},
                        {"status.defaultOutputSet", "Saída padrão definida como: {0}"},
                        {"status.defaultInputSet", "Entrada padrão definida como: {0}"},
                        {"status.volumeSet", "Volume de {0} definido para {1}%"},
                        {"status.volumeSetFailed", "Falha ao definir volume de {0}: {1}"},
                        {"status.systemSetFailed", "Falha ao definir dispositivo padrão do sistema: {0}"},
                        {"message.systemSetFailed", "Falha ao definir dispositivo padrão do sistema"},
                        {"status.vmConnected", "{0} conectado; {1}; DLL: {2}"},
                        {"status.modeAll", "todos os dispositivos"},
                        {"status.modePhysical", "visão de dispositivos físicos"},
                        {"status.vmReadFailed", "Falha ao conectar ou ler o Voicemeeter: {0}"},
                        {"status.vmInputSet", "Hardware Input 1 definido como: {0} - {1}"},
                        {"status.vmOutputSet", "A1 definido como: {0} - {1}"},
                        {"status.vmSetFailed", "Falha ao definir dispositivo do Voicemeeter: {0}"},
                        {"message.vmSetFailed", "Falha ao definir dispositivo do Voicemeeter"},
                        {"status.vmRestarted", "Comando para reiniciar o motor de áudio enviado ao Voicemeeter."},
                        {"status.vmRestartFailed", "Falha ao reiniciar o motor de áudio: {0}"}
                    }
                },
                {
                    "es", new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
                    {
                        {"tab.system", "Predeterminado del Sistema"},
                        {"tab.voicemeeter", "Voicemeeter"},
                        {"tab.language", "Idioma"},
                        {"button.refresh", "Actualizar"},
                        {"button.vmDefaults", "Aplicar valores de Voicemeeter"},
                        {"button.restart", "Reiniciar motor de audio"},
                        {"radio.physicalOnly", "Solo físicos/auriculares/integrados"},
                        {"radio.showAll", "Mostrar todos los dispositivos"},
                        {"group.output", "Reproducción / Salida"},
                        {"group.input", "Grabación / Entrada"},
                        {"group.volumeInput", "Volumen de Dispositivos de Entrada"},
                        {"group.volumeOutput", "Volumen de Dispositivos de Salida"},
                        {"group.vmInput", "Hardware Input 1 (Micrófono)"},
                        {"group.vmOutput", "A1 Hardware Out (Altavoces/Auriculares)"},
                        {"column.output", "Dispositivo de Salida"},
                        {"column.input", "Dispositivo de Entrada"},
                        {"column.driverInterface", "Controlador/Interfaz"},
                        {"column.status", "Estado"},
                        {"column.driver", "Controlador"},
                        {"column.hardwareId", "ID de hardware"},
                        {"language.title", "Idioma de la interfaz"},
                        {"language.description", "La interfaz se actualiza al instante y la opción se guarda para el próximo inicio."},
                        {"language.saved", "Idioma de la interfaz cambiado a: {0}"},
                        {"volume.none", "No hay dispositivos físicos con volumen ajustable"},
                        {"volume.unavailable", "No disponible"},
                        {"status.systemRead", "Se cargaron {0} dispositivos de salida y {1} dispositivos de entrada."},
                        {"status.systemReadFailed", "Error al leer los dispositivos de audio del sistema: {0}"},
                        {"status.startupDefaultsApplied", "Ruta predeterminada de Voicemeeter aplicada al iniciar: {0}"},
                        {"status.startupDefaultsMissing", "No se encontraron todos los dispositivos predeterminados de Voicemeeter al iniciar: {0}"},
                        {"status.vmDefaultOutput", "Salida predeterminada -> {0}"},
                        {"status.vmDefaultInput", "Entrada predeterminada -> {0}"},
                        {"status.vmDefaultOutputMissing", "No se encontró el dispositivo de salida predeterminado VoiceMeeter Input"},
                        {"status.vmDefaultInputMissing", "No se encontró el dispositivo de entrada predeterminado VoiceMeeter Output / B1"},
                        {"status.vmDefaultsFailed", "Error al aplicar la ruta predeterminada de Voicemeeter: {0}"},
                        {"status.defaultOutputSet", "Salida predeterminada establecida en: {0}"},
                        {"status.defaultInputSet", "Entrada predeterminada establecida en: {0}"},
                        {"status.volumeSet", "Volumen de {0} establecido en {1}%"},
                        {"status.volumeSetFailed", "Error al establecer el volumen de {0}: {1}"},
                        {"status.systemSetFailed", "Error al establecer el dispositivo predeterminado del sistema: {0}"},
                        {"message.systemSetFailed", "Error al establecer el dispositivo predeterminado del sistema"},
                        {"status.vmConnected", "{0} conectado; {1}; DLL: {2}"},
                        {"status.modeAll", "todos los dispositivos"},
                        {"status.modePhysical", "vista de dispositivos físicos"},
                        {"status.vmReadFailed", "Error al conectar o leer Voicemeeter: {0}"},
                        {"status.vmInputSet", "Hardware Input 1 establecido en: {0} - {1}"},
                        {"status.vmOutputSet", "A1 establecido en: {0} - {1}"},
                        {"status.vmSetFailed", "Error al establecer el dispositivo de Voicemeeter: {0}"},
                        {"message.vmSetFailed", "Error al establecer el dispositivo de Voicemeeter"},
                        {"status.vmRestarted", "Comando para reiniciar el motor de audio enviado a Voicemeeter."},
                        {"status.vmRestartFailed", "Error al reiniciar el motor de audio: {0}"}
                    }
                }
            };

        public static string NormalizeLanguage(string code)
        {
            if (string.IsNullOrWhiteSpace(code))
            {
                return DefaultLanguage;
            }

            string normalized = code.Trim().ToLowerInvariant();
            foreach (LanguageOption option in LanguageOptions)
            {
                if (string.Equals(option.Code, normalized, StringComparison.OrdinalIgnoreCase))
                {
                    return option.Code;
                }
            }

            return DefaultLanguage;
        }

        public static LanguageOption GetOption(string code)
        {
            string normalized = NormalizeLanguage(code);
            foreach (LanguageOption option in LanguageOptions)
            {
                if (string.Equals(option.Code, normalized, StringComparison.OrdinalIgnoreCase))
                {
                    return option;
                }
            }

            return LanguageOptions[0];
        }

        public static string T(string languageCode, string key)
        {
            Dictionary<string, string> table;
            if (!Strings.TryGetValue(NormalizeLanguage(languageCode), out table))
            {
                table = Strings[DefaultLanguage];
            }

            string value;
            if (table.TryGetValue(key, out value))
            {
                return value;
            }

            if (Strings[DefaultLanguage].TryGetValue(key, out value))
            {
                return value;
            }

            return key;
        }
    }
}
