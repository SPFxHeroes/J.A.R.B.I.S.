import { escape } from '@microsoft/sp-lodash-subset';
import { Version, DisplayMode } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneToggle
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';

import styles from './JarbisWebPart.module.scss';
import icons from './HeroIcons.module.scss';
//import * as strings from 'JarbisWebPartStrings';

import { IPowerItem } from './IPowerItem';
import { spfi, SPFx } from '@pnp/sp';
import '@pnp/sp/webs';
import '@pnp/sp/lists';
import '@pnp/sp/items';
import "@pnp/sp/items/get-all";
import { Caching } from "@pnp/queryable";

export interface IJarbisWebPartProps {
  name: string;
  primaryPower: string;
  secondaryPower: string;
  foregroundColor: string;
  backgroundColor: string;
  foregroundIcon: string;
  backgroundIcon: string;

  // The name of the SharePoint list that contains the powers
  list: string;

  // Indicates if the hero's powers should be shown at render time
  powersVisible: boolean;
}

export default class JarbisWebPart extends BaseClientSideWebPart<IJarbisWebPartProps> {

  private powers: IPowerItem[];

  public render(): void {
    const oldbuttons = this.domElement.getElementsByClassName(styles.generateButton);
    for (let b = 0; b < oldbuttons.length; b++) {
      oldbuttons[b].removeEventListener('click', this.onGenerateHero);
    }

    if (this.displayMode === DisplayMode.Edit && this.powers === undefined) {
      this.context.statusRenderer.displayLoadingIndicator(this.domElement, 'options');

      //load the powers
      this.getPowers().catch((error) => console.error(error));
      return;
    } else {
      this.context.statusRenderer.clearLoadingIndicator(this.domElement);
    }

    const hero = `
      <div class="${styles.logo} ${icons.heroIcons}">
        <i class="${this.getIconClass(escape(this.properties.backgroundIcon))} ${styles.background}" style="color:${escape(this.properties.backgroundColor)};"></i>
        <i class="${this.getIconClass(escape(this.properties.foregroundIcon))} ${styles.foreground}" style="color:${escape(this.properties.foregroundColor)};"></i>
      </div>
      <div class="${styles.name}">
        The ${escape(this.properties.name)}
      </div>`;

    const powers = `
      <div class="${styles.powers}">
        (${escape(this.properties.primaryPower)} + ${escape(this.properties.secondaryPower)})
      </div>`;

    const generateButton = `<button class=${styles.generateButton}>Generate</button>`;

    this.domElement.innerHTML = `
      <div class="${styles.jarbis}">
        ${hero}
        ${this.properties.powersVisible ? powers : ""}
        ${this.displayMode === DisplayMode.Edit ? generateButton : ""}
      </div>`;

    const buttons = this.domElement.getElementsByClassName(styles.generateButton);
    for (let b = 0; b < buttons.length; b++) {
      buttons[b].addEventListener('click', this.onGenerateHero);
    }
  }

  /**
  * Gets the list of powers from SharePoint
  *
  * @private
  * @memberof JarbisWebPart
  */
  private getPowers = async (): Promise<void> => {
    const sp = spfi().using(SPFx(this.context));

    // Get the list of powers from SharePoint using the name of the library specified in the property pane
    this.powers = await sp.web.lists.getByTitle(this.properties.list).items.select('Title', 'Icon', 'Colors', 'Prefix', 'Main').using(Caching()).getAll();

    // Re-render the web part
    this.render();
  }

  /**
  * Generates a new hero with random values
  *
  * @param {MouseEvent} _event Unused event parameter
  * @memberof JarbisWebPart
     */
  public onGenerateHero = (_event: MouseEvent): void => {
    // Get a random power (list item) from the list of powers
    const power1: IPowerItem = this.getRandomItem(this.powers);

    // Get another random power (list item) from the list of powers, excluding the first power
    const power2: IPowerItem = this.getRandomItem(this.powers, power1);

    // Get the titles from each of the powers and save them to our properties
    this.properties.primaryPower = power1.Title;
    this.properties.secondaryPower = power2.Title;

    // Get a random color for the background choosing from the combined color suggestions for the two powers
    this.properties.backgroundColor = this.getRandomItem([...power1.Colors, ...power2.Colors]);
    // Get a random color for the foreground choosing from the same list of suggestions but excluding the background color
    this.properties.foregroundColor = this.getRandomItem([...power1.Colors, ...power2.Colors], this.properties.backgroundColor);

    // Get a random icon for the background choosing from a fixed list of background icons
    this.properties.backgroundIcon = this.getRandomItem(['StarburstSolid', 'CircleShapeSolid', 'HeartFill', 'SquareShapeSolid', 'ShieldSolid']);
    // Get a random icon for the foreground choosing from the combined icon suggestions for the two powers
    this.properties.foregroundIcon = this.getRandomItem([...power1.Icon, ...power2.Icon], this.properties.backgroundIcon);

    // Get the prefix choosing from the combined prefix suggestions for the two powers
    const prefix = this.getRandomItem([...power1.Prefix, ...power2.Prefix]);
    // Get the main portion of the name by choosing from the combined main suggestions
    //  for the two powers excluding the prefix since there is some overlap
    const main = this.getRandomItem([...power1.Main, ...power2.Main], prefix);

    // Store the name of the hero by combining the prefix with the main
    this.properties.name = prefix + ' ' + main;

    // Re-render the web part
    this.render();
  }

  /**
  * Gets a random value from an array of choices, excluding a specific value
  *
  * @private
  * @param {any[]} choices The array of choices to pick from
  * @param {any} exclusion The value to exclude from the choices
  * @memberof JarbisWebPart
  */

  private getRandomItem = (choices: any[], exclusion?: any): any => {
    // Filter the choices to exclude the previous value
    const filteredChoices = choices.filter((value) => value !== exclusion);

    // If there are any choices left, pick a random one
    if (filteredChoices.length) {
      return filteredChoices[Math.floor(Math.random() * filteredChoices.length)];
    }

    // Otherwise, return an empty string
    return "";
  }

  private getIconClass(iconName: string): string {
    const iconKey: string = "icon" + iconName;
    if (this.hasKey(icons, iconKey)) {
      return icons[iconKey];
    }
  }

  private hasKey<O extends object>(obj: O, key: PropertyKey): key is keyof O {
    return key in obj;
  }

  protected onThemeChanged(currentTheme: IReadonlyTheme | undefined): void {
    if (!currentTheme) {
      return;
    }

    const {
      semanticColors
    } = currentTheme;

    if (semanticColors) {
      this.domElement.style.setProperty('--bodyText', semanticColors.bodyText || null);
      this.domElement.style.setProperty('--link', semanticColors.link || null);
      this.domElement.style.setProperty('--linkHovered', semanticColors.linkHovered || null);
    }

  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          groups: [
            {
              groupFields: [
                PropertyPaneToggle('powersVisible', {
                  label: "Powers",
                  onText: "Visible",
                  offText: "Hidden"
                })
              ]
            }
          ]
        }
      ]
    };
  }

  protected onDispose(): void {
    const oldbuttons = this.domElement.getElementsByClassName(styles.generateButton);
    for (let b = 0; b < oldbuttons.length; b++) {
      oldbuttons[b].removeEventListener('click', this.onGenerateHero);
    }
  }
}
