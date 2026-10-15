/** An error the user can understand; the conversion is aborted and nothing is written. */
export class ConversionError extends Error {
  constructor(message: string) {
    super(message);
    this.name = 'ConversionError';
  }
}
