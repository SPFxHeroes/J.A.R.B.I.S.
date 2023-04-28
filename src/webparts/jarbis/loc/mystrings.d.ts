declare interface IJarbisWebPartStrings {

  /**
   * The label for the generate button
   */
  GenerateButtonLabel: string;

  /**
   * The label for the "Show Powers" property field in the web part properties
   */
  ShowPowersFieldLabel: string;

  /*
   * The label for the "Show Powers" toggle when it is off
   */
  ShowPowersToggleOffText: string;

  /*
   * The label for the "Show Powers" toggle when it is off
   */
  ShowPowersToggleOnText: string;

  /**
    * The description for the hero (e.g. The Mighty Coder)
    */
  HeroDescription: string;
}

declare module 'JarbisWebPartStrings' {
  const strings: IJarbisWebPartStrings;
  export = strings;
}