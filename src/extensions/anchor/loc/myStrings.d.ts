declare interface IAnchorCommandSetStrings {
  Command1: string;
  Command2: string;
}

declare module 'AnchorCommandSetStrings' {
  const strings: IAnchorCommandSetStrings;
  export = strings;
}
