import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version, DisplayMode } from '@microsoft/sp-core-library';
import { SPComponentLoader } from '@microsoft/sp-loader';
import { BaseClientSideWebPart } from "@microsoft/sp-webpart-base";
import { IPropertyPaneConfiguration, IPropertyPaneField, PropertyPaneTextField, PropertyPaneToggle } from "@microsoft/sp-property-pane";
import { IScriptEditorProps } from './components/IScriptEditorProps';
import { IScriptEditorWebPartProps } from './IScriptEditorWebPartProps';
import PropertyPaneLogo from './PropertyPaneLogo';
import { PrincipalType, PropertyFieldNumber, PropertyFieldPeoplePicker } from '@pnp/spfx-property-controls';
import { UserGroupCheck } from './AccessCheck';
import { sp } from '@pnp/sp/presets/all';

export default class ScriptEditorWebPart extends BaseClientSideWebPart<IScriptEditorWebPartProps> {
    public _propertyPaneHelper;
    private _unqiueId;
    private _externalScriptContent;

    constructor() {
        super();
        this.scriptUpdate = this.scriptUpdate.bind(this);
    }

    public scriptUpdate(_property: string, _oldVal: string, newVal: string) {
        this.properties.script = newVal;
        this._propertyPaneHelper.initialValue = newVal;
    }

    protected async onInit(): Promise<void> {
        if (this.properties.useExternalScript) {
            try {
                const prefix = this.properties.externalScript.indexOf('?') === -1 ? '?' : '&';
                const response = await fetch(`${this.properties.externalScript}${prefix}pnp=${new Date().getTime()}`);
                this._externalScriptContent = await response.text();
            } catch {
                this._externalScriptContent = 'Failed to load external script.';
            }
        }

        await super.onInit();
        sp.setup({
            spfxContext: this.context as any
        });
    }

    public async render(): Promise<void> {
        this._unqiueId = this.context.instanceId;
        if (this.displayMode == DisplayMode.Read) {
            if (this.properties.removePadding) {
                let element = this.domElement.parentElement;
                // check up to 5 levels up for padding and exit once found
                for (let i = 0; i < 5; i++) {
                    const style = window.getComputedStyle(element);
                    const hasPadding = style.paddingTop !== "0px";
                    if (hasPadding) {
                        element.style.paddingTop = "0px";
                        element.style.paddingBottom = "0px";
                        element.style.marginTop = "0px";
                        element.style.marginBottom = "0px";
                    }
                    element = element.parentElement;
                }
            }

            ReactDom.unmountComponentAtNode(this.domElement);
            if (this.properties.audiences) {
                const checker = new UserGroupCheck(this.properties.audiences, this.properties.audienceCacheDuration, this.context);
                const isInAudience = await checker.CheckAudiences();
                if (!isInAudience) return;
            }

            this.domElement.innerHTML = this.properties.useExternalScript ? this._externalScriptContent : this.properties.script;
            this.executeScript(this.domElement);
        } else {
            this.renderEditor();
        }
    }

    private async renderEditor() {
        // Dynamically load the editor pane to reduce overall bundle size
        const editorPopUp = await import(
            /* webpackChunkName: 'scripteditor' */
            './components/ScriptEditor'
        );
        const element: React.ReactElement<IScriptEditorProps> = React.createElement(
            editorPopUp.default,
            {
                script: this.properties.script,
                title: this.properties.title,
                propPaneHandle: this.context.propertyPane,
                key: "pnp" + new Date().getTime()
            }
        );
        ReactDom.render(element, this.domElement);
    }

    protected get dataVersion(): Version {
        return Version.parse('1.0');
    }

    protected async loadPropertyPaneResources(): Promise<void> {
        //import { PropertyFieldCodeEditor, PropertyFieldCodeEditorLanguages } from '@pnp/spfx-property-controls/lib/PropertyFieldCodeEditor';
        const editorProp = await import(
            /* webpackChunkName: 'scripteditor' */
            '@pnp/spfx-property-controls/lib/PropertyFieldCodeEditor'
        );

        this._propertyPaneHelper = editorProp.PropertyFieldCodeEditor('scriptCode', {
            label: 'Edit HTML Code',
            panelTitle: 'Edit HTML Code',
            initialValue: this.properties.script,
            onPropertyChange: this.scriptUpdate,
            properties: this.properties,
            disabled: false,
            key: 'codeEditorFieldId',
            language: editorProp.PropertyFieldCodeEditorLanguages.HTML
        });
    }

    protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
        const webPartOptions: IPropertyPaneField<any>[] = [
            PropertyPaneTextField("title", {
                label: "Title to show in edit mode",
                value: this.properties.title
            }),
            PropertyPaneToggle("removePadding", {
                label: "Remove top/bottom padding of web part container",
                checked: this.properties.removePadding,
                onText: "Remove padding",
                offText: "Keep padding"
            }),
            PropertyPaneToggle("spPageContextInfo", {
                label: "Enable classic _spPageContextInfo",
                checked: this.properties.spPageContextInfo,
                onText: "Enabled",
                offText: "Disabled"
            }),
            PropertyPaneToggle('useExternalScript', {
                label: 'Use script from an external URL instead of inline script',
                onText: "Use external script",
                offText: "Use inline script"
            }),
            PropertyFieldPeoplePicker('audiences', {
                label: 'Target Audience',
                initialData: this.properties.audiences,
                allowDuplicate: false,
                principalType: [PrincipalType.SharePoint, PrincipalType.Users, PrincipalType.Security],
                onPropertyChange: this.onPropertyPaneFieldChanged,
                context: this.context as any,
                properties: this.properties,
                onGetErrorMessage: null,
                deferredValidationTime: 0,
                key: 'audienceTargeting'
            }),
            PropertyFieldNumber('audienceCacheDuration', {
                key: 'audienceCacheDuration',
                label: 'Audience cache duration in hours',
                description: 'Duration in hours to cache the audience information',
                value: this.properties.audienceCacheDuration || 24,
                minValue: 1,
            })
        ];

        if (this.properties.useExternalScript) {
            webPartOptions.push(PropertyPaneTextField("externalScript", {
                label: "Script URL",
                value: this.properties.externalScript,
            }));
        } else {
            webPartOptions.push(this._propertyPaneHelper);
        }

        if (this.context.sdks.microsoftTeams) {
            const config = PropertyPaneToggle("teamsContext", {
                label: "Enable teams context as _teamsContexInfo",
                checked: this.properties.teamsContext,
                onText: "Enabled",
                offText: "Disabled"
            });
            webPartOptions.push(config);
        }
        webPartOptions.push(new PropertyPaneLogo());

        return {
            pages: [
                {
                    groups: [
                        {
                            groupFields: webPartOptions
                        }
                    ]
                }
            ]
        };
    }


    // SharePoint Online CSP allows 'unsafe-eval' but not 'unsafe-inline', so inline code
    // must be compiled with the Function constructor instead of injected <script> tags.
    private runCode(code: string, thisArg: any, source: string, event?: Event): void {
        if (!code || code.trim().length === 0) return;
        try {
            const func = new Function("event", code);
            func.call(thisArg, event);
        } catch (error) {
            if (console.error) {
                if (error instanceof EvalError) {
                    console.error(`Script Editor: ${source} blocked by Content Security Policy ('unsafe-eval' not allowed). Move the code to an external .js file and add its host to Trusted script sources in the SharePoint admin center.`, error);
                } else {
                    console.error(`Script Editor: failed to execute ${source}:`, error);
                }
            }
        }
    }

    private evalScript(elem): void {
        // External scripts are handled in executeScript method
        if (elem.src && elem.src.length > 0) return;

        const data = (elem.text || elem.textContent || elem.innerHTML || "");
        this.runCode(data, window, "inline script");
    }

    // Inline event handler attributes (onclick="..." etc.) and javascript: URLs are blocked by CSP
    // ('unsafe-inline'), so convert them to listeners compiled with the Function constructor.
    private rebindInlineHandlers(element: HTMLElement): void {
        const nodes = element.querySelectorAll("*");
        for (let i = 0; i < nodes.length; i++) {
            const node = nodes[i] as HTMLElement;
            const attributes = Array.prototype.slice.call(node.attributes) as Attr[];
            for (const attr of attributes) {
                const name = attr.name.toLowerCase();
                if (name.indexOf("on") === 0 && name.length > 2) {
                    const code = attr.value;
                    node.removeAttribute(attr.name);
                    node.addEventListener(name.substring(2), (ev: Event) => this.runCode(code, node, `${name} handler`, ev));
                } else if ((name === "href" || name === "src" || name === "action") && /^\s*javascript:/i.test(attr.value)) {
                    const code = attr.value.replace(/^\s*javascript:/i, "");
                    node.setAttribute(attr.name, name === "href" ? "#" : "");
                    node.addEventListener(name === "action" ? "submit" : "click", (ev: Event) => {
                        ev.preventDefault();
                        this.runCode(decodeURIComponent(code), node, "javascript: URL", ev);
                    });
                }
            }
        }
    }

    private nodeName(elem, name) {
        return elem.nodeName && elem.nodeName.toUpperCase() === name.toUpperCase();
    }

    // Finds and executes scripts in a newly added element's body.
    // Needed since innerHTML does not run scripts.
    //
    // Argument element is an element in the dom.
    private async executeScript(element: HTMLElement) {

        if (this.properties.spPageContextInfo && !window["_spPageContextInfo"]) {
            window["_spPageContextInfo"] = this.context.pageContext.legacyPageContext;
        }

        if (this.properties.teamsContext && !window["_teamsContexInfo"]) {
            window["_teamsContexInfo"] = this.context.sdks.microsoftTeams.context;
        }

        // Define global name to tack scripts on in case script to be loaded is not AMD/UMD
        (<any>window).ScriptGlobal = {};

        // main section of function
        const scripts = [];
        const children_nodes = element.getElementsByTagName("script");

        for (let i = 0; children_nodes[i]; i++) {
            const child: any = children_nodes[i];
            const type = (child.type || "").toLowerCase();
            if (!type || type === "text/javascript" || type === "application/javascript") {
                scripts.push(child);
            }
        }

        const urls = [];
        const onLoads: string[] = [];
        for (let i = 0; scripts[i]; i++) {
            const scriptTag = scripts[i];
            if (scriptTag.src && scriptTag.src.length > 0) {
                urls.push(scriptTag.src);
            }
            // onload="..." is an inline handler that CSP blocks, so read the attribute text instead
            const onLoad = scriptTag.getAttribute("onload");
            if (onLoad && onLoad.trim().length > 0) {
                onLoads.push(onLoad);
            }
        }

        let oldamd = null;
        if (window["define"] && window["define"].amd) {
            oldamd = window["define"].amd;
            window["define"].amd = null;
        }

        for (let i = 0; i < urls.length; i++) {
            try {
                let scriptUrl = urls[i];
                // Add unique param to force load on each run to overcome smart navigation in the browser as needed
                const prefix = scriptUrl.indexOf('?') === -1 ? '?' : '&';
                scriptUrl += prefix + 'pnp=' + new Date().getTime();
                await SPComponentLoader.loadScript(scriptUrl, { globalExportsName: "ScriptGlobal" });
            } catch (error) {
                if (console.error) {
                    console.error(`Script Editor: failed to load ${urls[i]}. If the console shows a Content Security Policy violation, add this script's host to Trusted script sources in the SharePoint admin center.`, error);
                }
            }
        }
        if (oldamd) {
            window["define"].amd = oldamd;
        }

        for (let i = 0; scripts[i]; i++) {
            const scriptTag = scripts[i];
            if (scriptTag.parentNode) { scriptTag.parentNode.removeChild(scriptTag); }
            this.evalScript(scripts[i]);
        }
        this.rebindInlineHandlers(element);
        // execute any onload people have added
        for (let i = 0; onLoads[i]; i++) {
            this.runCode(onLoads[i], window, "script onload handler");
        }
    }
}
