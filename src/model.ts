import { DocumentRegistry } from '@jupyterlab/docregistry';
import { PartialJSONObject } from '@lumino/coreutils';
import { ISignal, Signal } from '@lumino/signaling';

/**
 * A minimal document model for binary documents that shouldn't be loaded as text.
 * This model does NOT load file content, preventing UTF-8 encoding errors.
 */
export class DocModel implements DocumentRegistry.IModel {
  /**
   * The default kernel name
   */
  readonly defaultKernelName = '';

  /**
   * The default kernel language
   */
  readonly defaultKernelLanguage = '';

  /**
   * Whether the model is collaborative
   */
  readonly collaborative = false;

  /**
   * A signal emitted when the content changes (never emitted for binary files)
   */
  readonly contentChanged: ISignal<this, void> = new Signal<this, void>(this);

  /**
   * A signal emitted when the state changes (never emitted for binary files)
   */
  readonly stateChanged: ISignal<this, any> = new Signal<this, any>(this);

  /**
   * The shared model (null for non-collaborative documents)
   */
  readonly sharedModel: any = null;

  /**
   * The dirty state of the model
   */
  get dirty(): boolean {
    return false;
  }

  /**
   * The read only state of the model
   */
  get readOnly(): boolean {
    return true;
  }

  /**
   * Serialize the model to a string (not used for binary files)
   */
  toString(): string {
    return '';
  }

  /**
   * Deserialize the model from a string (not used for binary files)
   */
  fromString(value: string): void {
    // Do nothing - we don't load content for binary files
  }

  /**
   * Serialize the model to JSON (not used for binary files)
   */
  toJSON(): PartialJSONObject {
    return {};
  }

  /**
   * Deserialize the model from JSON (not used for binary files)
   */
  fromJSON(value: PartialJSONObject): void {
    // Do nothing - we don't load content for binary files
  }

  /**
   * Initialize the model (no-op)
   */
  initialize(): void {
    // Do nothing
  }

  /**
   * Dispose of the model
   */
  dispose(): void {
    // Do nothing
  }

  /**
   * Whether the model is disposed
   */
  get isDisposed(): boolean {
    return false;
  }
}
