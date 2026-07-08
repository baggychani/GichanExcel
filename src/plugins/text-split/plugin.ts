import {
  ICommandService,
  Inject,
  Injector,
  Plugin,
  UniverInstanceType,
} from "@univerjs/core";
import { IMenuManagerService } from "@univerjs/ui";
import { OpenTextSplitDialogCommand } from "./command";
import { textSplitMenuSchema } from "./schema";

/**
 * Univer's DI container (redi) resolves `createInstance(plugin, options)` by
 * filling *undecorated* constructor params positionally with `options` first,
 * then resolving *decorated* params (e.g. `@Inject(Injector)`) from the
 * container. Without the decorator, `injector` below would silently receive
 * `undefined` instead of the real Injector — which is why the menu never
 * registered before.
 */
export class GichanTextSplitPlugin extends Plugin {
  static override pluginName = "GICHAN_TEXT_SPLIT_PLUGIN";
  static override type = UniverInstanceType.UNIVER_SHEET;

  protected override _injector: Injector;

  constructor(_config: unknown, @Inject(Injector) injector: Injector) {
    super();
    this._injector = injector;
  }

  override onStarting(): void {
    this._injector
      .get(ICommandService)
      .registerCommand(OpenTextSplitDialogCommand);
  }

  override onRendered(): void {
    this._injector.get(IMenuManagerService).mergeMenu(textSplitMenuSchema);
  }
}

