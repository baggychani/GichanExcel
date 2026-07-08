import {
  LogLevel,
  Univer,
  type DependencyOverride,
  type IUniverConfig,
  type Plugin,
  type PluginCtor,
} from "@univerjs/core";
import { FUniver } from "@univerjs/core/facade";

interface Preset {
  plugins: Array<
    PluginCtor<Plugin> | [PluginCtor<Plugin>, ConstructorParameters<PluginCtor<Plugin>>[0]]
  >;
}

type PluginEntry =
  | PluginCtor<Plugin>
  | [PluginCtor<Plugin>, ConstructorParameters<PluginCtor<Plugin>>[0]];

interface CreateUniverOptions extends Partial<IUniverConfig> {
  presets: Array<Preset | [Preset, unknown]>;
  plugins?: PluginEntry[];
  override?: DependencyOverride;
}

export function createUniver(options: CreateUniverOptions) {
  const { presets, plugins, override = [], ...config } = options;
  const univer = new Univer({
    logLevel: LogLevel.WARN,
    ...config,
    override,
  });
  const registeredPlugins = new Map<
    string,
    { plugin: PluginCtor<Plugin>; options?: ConstructorParameters<PluginCtor<Plugin>>[0] }
  >();

  presets.forEach((presetEntry) => {
    const preset = Array.isArray(presetEntry) ? presetEntry[0] : presetEntry;
    preset.plugins.forEach((entry) => {
      const [plugin, pluginOptions] = Array.isArray(entry) ? [entry[0], entry[1]] : [entry];
      registeredPlugins.delete(plugin.pluginName);
      registeredPlugins.set(plugin.pluginName, { plugin, options: pluginOptions });
    });
  });

  plugins?.forEach((entry) => {
    const [plugin, pluginOptions] = Array.isArray(entry) ? [entry[0], entry[1]] : [entry];
    if (registeredPlugins.has(plugin.pluginName)) {
      throw new Error(`Plugin ${plugin.pluginName} is already registered.`);
    }
    registeredPlugins.set(plugin.pluginName, { plugin, options: pluginOptions });
  });

  registeredPlugins.forEach(({ plugin, options: pluginOptions }) => {
    univer.registerPlugin(plugin, pluginOptions);
  });

  return {
    univer,
    univerAPI: FUniver.newAPI(univer),
  };
}
