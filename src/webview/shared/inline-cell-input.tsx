import { onMount } from "solid-js";

export function InlineCellInput(props: {
    class: string;
    value: string;
    dataRole?: string;
    onUpdateDraft: (value: string) => void;
    onCommit: () => void;
    onCancel: () => void;
}) {
    let inputElement: HTMLInputElement | undefined;

    onMount(() => {
        inputElement?.focus();
        inputElement?.select();
    });

    return (
        <input
            ref={(element) => {
                inputElement = element;
            }}
            class={props.class}
            data-role={props.dataRole}
            type="text"
            value={props.value}
            onBlur={() => {
                setTimeout(() => props.onCommit(), 0);
            }}
            onInput={(event) => props.onUpdateDraft(event.currentTarget.value)}
            onClick={(event) => event.stopPropagation()}
            onDblClick={(event) => event.stopPropagation()}
            onKeyDown={(event) => {
                if (event.key === "Enter" || event.key === "Tab") {
                    event.preventDefault();
                    props.onCommit();
                } else if (event.key === "Escape") {
                    event.preventDefault();
                    props.onCancel();
                }
            }}
        />
    );
}
