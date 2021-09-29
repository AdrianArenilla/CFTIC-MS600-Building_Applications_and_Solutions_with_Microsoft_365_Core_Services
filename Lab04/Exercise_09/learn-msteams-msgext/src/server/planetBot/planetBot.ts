import {
    TeamsActivityHandler,
    TurnContext,
    MessageFactory,
    CardFactory, MessagingExtensionAction, MessagingExtensionActionResponse, MessagingExtensionAttachment,
    MessagingExtensionQuery, MessagingExtensionResponse
} from "botbuilder";

import * as Util from "util";
import * as debug from "debug";
import { find, sortBy } from "lodash";

const TextEncoder = Util.TextEncoder;
const log = debug("msteams");

export class PlanetBot extends TeamsActivityHandler {

    protected handleTeamsMessagingExtensionFetchTask(context: TurnContext, action: MessagingExtensionAction): Promise<MessagingExtensionActionResponse> {

        // load planets & sort them by their order from the sun
        const planets: any = require("./planets.json");
        const sortedPlanets: any = sortBy(planets, ["id"])
            .map((planet) => {
                return { value: planet.id, title: planet.name };
            });

        // load card template
        const adaptiveCardSource: any = require("./planetSelectorCard.json");
        // locate the planet selector
        const planetChoiceSet: any = find(adaptiveCardSource.body, { id: "planetSelector" });
        // update choice set with planets
        planetChoiceSet.choices = sortedPlanets;
        // load the adaptive card
        const adaptiveCard = CardFactory.adaptiveCard(adaptiveCardSource);

        const response: MessagingExtensionActionResponse = {
            task: {
                type: "continue",
                value: {
                    card: adaptiveCard,
                    title: "Planet Selector",
                    height: 150,
                    width: 500
                }
            }
        } as MessagingExtensionActionResponse;

        return Promise.resolve(response);
    }

    private getPlanetDetailCard(selectedPlanet: any): MessagingExtensionAttachment {
        // load display card
        const adaptiveCardSource: any = require("./planetDisplayCard.json");

        // update planet fields in display card
        adaptiveCardSource.actions[0].url = selectedPlanet.wikiLink;
        find(adaptiveCardSource.body, { id: "cardHeader" }).items[0].text = selectedPlanet.name;
        const cardBody: any = find(adaptiveCardSource.body, { id: "cardBody" });
        find(cardBody.items, { id: "planetSummary" }).text = selectedPlanet.summary;
        find(cardBody.items, { id: "imageAttribution" }).text = "*Image attribution: " + selectedPlanet.imageAlt + "*";
        const cardDetails: any = find(cardBody.items, { id: "planetDetails" });
        cardDetails.columns[0].items[0].url = selectedPlanet.imageLink;
        find(cardDetails.columns[1].items[0].facts, { id: "orderFromSun" }).value = selectedPlanet.id;
        find(cardDetails.columns[1].items[0].facts, { id: "planetNumSatellites" }).value = selectedPlanet.numSatellites;
        find(cardDetails.columns[1].items[0].facts, { id: "solarOrbitYears" }).value = selectedPlanet.solarOrbitYears;
        find(cardDetails.columns[1].items[0].facts, { id: "solarOrbitAvgDistanceKm" }).value = Number(selectedPlanet.solarOrbitAvgDistanceKm).toLocaleString();

        // return the adaptive card
        return CardFactory.adaptiveCard(adaptiveCardSource);
    }

    protected handleTeamsMessagingExtensionSubmitAction(context: TurnContext, action: MessagingExtensionAction): Promise<MessagingExtensionActionResponse> {
        switch (action.commandId) {
            case "planetExpanderAction":
                // load planets
                const planets: any = require("./planets.json");
                // get the selected planet
                const selectedPlanet: any = planets.filter((planet) => planet.id === action.data.planetSelector)[0];
                const adaptiveCard = this.getPlanetDetailCard(selectedPlanet);

                // generate the response
                return Promise.resolve({
                    composeExtension: {
                        type: "result",
                        attachmentLayout: "list",
                        attachments: [adaptiveCard]
                    }
                } as MessagingExtensionActionResponse);
            default:
                throw new Error("NotImplemented");
        }
    }

    private getPlanetResultCard(selectedPlanet: any): MessagingExtensionAttachment {
        return CardFactory.heroCard(selectedPlanet.name, selectedPlanet.summary, [selectedPlanet.imageLink]);
    }

    protected handleTeamsMessagingExtensionQuery(context: TurnContext, query: MessagingExtensionQuery): Promise<MessagingExtensionResponse> {
        // get the search query
        let searchQuery = "";
        if (query && query.parameters && query.parameters[0].name === "searchKeyword" && query.parameters[0].value) {
            searchQuery = query.parameters[0].value.trim().toLowerCase();
        }

        // load planets
        const planets: any = require("./planets.json");
        // search results
        let queryResults: string[] = [];

        switch (searchQuery) {
            case "inner":
                // get all planets inside asteroid belt
                queryResults = planets.filter((planet) => planet.id <= 4);
                break;
            case "outer":
                // get all planets outside asteroid belt
                queryResults = planets.filter((planet) => planet.id > 4);
                break;
            default:
                // get the specified planet
                queryResults.push(planets.filter((planet) => planet.name.toLowerCase() === searchQuery)[0]);
        }

        // get the results as cards
        const searchResultsCards: MessagingExtensionAttachment[] = [];
        queryResults.forEach((planet) => {
            searchResultsCards.push(this.getPlanetResultCard(planet));
        });

        const response: MessagingExtensionResponse = {
            composeExtension: {
                type: "result",
                attachmentLayout: "list",
                attachments: searchResultsCards
            }
        } as MessagingExtensionResponse;

        return Promise.resolve(response);
    }
}
