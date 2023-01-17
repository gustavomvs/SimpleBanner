import * as React from "react";
import styles from "./SimpleBanner.module.scss";
import { ISimpleBannerProps } from "./ISimpleBannerProps";
import { spfi, SPFx as spSPFx } from "@pnp/sp";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import "@pnp/sp/attachments";
import "@pnp/sp/folders";
import "@pnp/sp/files";
import {
  FilePicker,
  IFilePickerResult,
} from "@pnp/spfx-controls-react/lib/FilePicker";
import { useBoolean } from "@fluentui/react-hooks";
import { useState, useEffect } from "react";
import {
  ILabelStyleProps,
  ITextFieldStyles,
  ITooltipHostStyles,
  TextField,
} from "office-ui-fabric-react";
import {
  DefaultButton,
  PrimaryButton,
  Dropdown,
  IDropdownStyles,
  IDropdownOption,
  Toggle,
  CommandButton,
  Panel,
  ILabelStyles,
} from "@fluentui/react";
import * as strings from "SimpleBannerWebPartStrings";

const textFieldStyles: Partial<ITextFieldStyles> = {
  fieldGroup: { marginBottom: "5px", fontSize: "1rem" },
  subComponentStyles: { label: getLabelStyles },
};

const hostStyles: Partial<ITooltipHostStyles> = {
  root: { height: "20px", margin: "8px", fontSize: "1rem" },
};

const dropdownStyles: Partial<IDropdownStyles> = {
  dropdown: { fontSize: "1rem" },
  label: { fontSize: "0.9rem" },
};

const toggleStyles: Partial<IDropdownStyles> = {
  label: { fontSize: "0.9rem" },
};

const options: IDropdownOption[] = [
  { key: "100%", text: "100%" },
  { key: "75%", text: "75%" },
  { key: "50%", text: "50%" },
  { key: "25%", text: "25%" },
];

interface SimpleBanner {
  urlImage: string;
  itemID: number;
  urlDestiny: string;
  newAbe: boolean;
  size: string;
  alt: string;
}

interface SimpleBannerTemp {
  urlImage: string;
  newAbe: boolean;
  urlDestiny: string;
  size: string;
  alt: string;
}

function getLabelStyles(props: ILabelStyleProps): ILabelStyles {
  return {
    root: {
      fontSize: "0.9rem",
    },
  };
}

const Simplebanner: React.FunctionComponent<ISimpleBannerProps> = (props) => {
  const sp = spfi().using(spSPFx(props.context));

  const simpleDefault: SimpleBanner = {
    urlImage: null,
    itemID: props.itemId,
    urlDestiny: "",
    newAbe: true,
    size: "100%",
    alt: "",
  };

  const [isOpen, { setTrue: openPanel, setFalse: dismissPanel }] =
    useBoolean(false);

  const [simple, setSimple] = useState<SimpleBanner>(simpleDefault);

  const [simpleTemp, setSimpleTemp] = useState<SimpleBannerTemp>(simpleDefault);

  const [file, setFile] = useState<IFilePickerResult>();
  const [preview, setPreview] = useState("");

  const meuInit = async (): Promise<void> => {
    if (simple.itemID) {
      await sp.web.lists
        .getByTitle("SimpleBanners")
        .items.getById(simple.itemID)
        .select(
          "Id",
          "FileRef",
          "newAbe",
          "size",
          "urlDestiny",
          "FileLeafRef"
        )()
        .then((res) => {
          setSimple({
            urlImage: res.FileRef,
            itemID: res.Id,
            urlDestiny: res.urlDestiny,
            newAbe: res.newAbe,
            size: res.size,
            alt: res.alt,
          });
        });
    }
  };

  useEffect(() => {
    // eslint-disable-next-line @typescript-eslint/no-floating-promises
    meuInit();
  }, []);

  const changesize = (
    event: React.FormEvent<HTMLDivElement>,
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    option?: any,
    index?: number
  ): void => {
    setSimple((old) => ({
      ...old,
      size: option.text,
    }));
  };

  const changeurlDestiny = (
    event: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>,
    newValue?: string
  ): void => {
    setSimple((old) => ({
      ...old,
      urlDestiny: newValue,
    }));
  };

  function changenewAbe(
    _ev: React.MouseEvent<HTMLElement>,
    checked?: boolean
  ): void {
    setSimple((old) => ({
      ...old,
      newAbe: checked,
    }));
  }

  const changeAlt = (
    event: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>,
    newValue?: string
  ): void => {
    setSimple((old) => ({
      ...old,
      alt: newValue,
    }));
  };

  const changeImg = (file: IFilePickerResult[]): void => {
    const oFile = file[0];
    if (oFile.fileAbsoluteUrl) {
      setPreview(oFile.fileAbsoluteUrl);
    } else {
      setPreview(oFile.previewDataUrl);
    }
    setSimple((old) => ({
      ...old,
      urlImage: oFile.previewDataUrl
        ? oFile.previewDataUrl
        : oFile.fileAbsoluteUrl,
    }));
    setFile(oFile);
  };

  const Save = async (): Promise<void> => {
    if (preview) {
      const fileResultContent = await file.downloadFileContent();
      const result = await sp.web
        .getFolderByServerRelativePath("SimpleBanners")
        .files.addUsingPath(fileResultContent.name, fileResultContent, {
          Overwrite: true,
        });

      const item = await result.file.getItem();

      await item.update({
        newAbe: simple.newAbe,
        urlDestiny: simple.urlDestiny,
        size: simple.size,
        alt: simple.alt,
      });

      await item
        .select("Id", "FileLeafRef")()
        .then((res) => {
          const nameSplit = res.FileLeafRef.split(".");
          const nameFinal = nameSplit[nameSplit.length - 1];
          const nameInitial = res.FileLeafRef.slice(
            0,
            res.FileLeafRef.length - nameFinal.length - 1
          ).slice(0, 200);
          // eslint-disable-next-line @typescript-eslint/no-floating-promises
          sp.web.lists
            .getByTitle("SimpleBanners")
            .items.getById(res.Id)
            .update({
              FileLeafRef: `${nameInitial}_${res.Id}_`,
            })
            .then(() => {
              setSimple((old) => ({
                ...old,
                itemID: res.Id,
              }));
            });

          props.updatePropety(res.Id);
          props.updateFileName(file.fileName);
          props.updateFileSize(file.fileSize);
        });
    } else {
      if (simple.itemID) {
        await sp.web.lists
          .getByTitle("SimpleBanners")
          .items.getById(simple.itemID)
          .update({
            newAbe: simple.newAbe === false ? false : true,
            urlDestiny: simple.urlDestiny,
            size: simple.size,
          })
          .then((res) => {
            setSimple((old) => ({
              ...old,
              urlDestiny: simple.urlDestiny,
              newAbe: simple.newAbe,
              size: simple.size,
              alt: simple.alt,
            }));
          });
      }
    }
    dismissPanel();
    setPreview("");
  };

  const onClickBanner = (): void => {
    if (simple.urlDestiny) {
      window.open(simple.urlDestiny, simple.newAbe ? "_blank" : "_self");
    }
  };

  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  const onRenderFooterContent = (): any => (
    <div className={styles.footerPanel}>
      <DefaultButton
        onClick={() => {
          // eslint-disable-next-line @typescript-eslint/no-floating-promises
          Save();
        }}
      >
        {strings.save}
      </DefaultButton>
      <PrimaryButton
        onClick={() => {
          dismissPanel();
          setPreview("");
          setSimple((old) => ({
            ...old,
            urlImage: simpleTemp.urlImage,
            urlDestiny: simpleTemp.urlDestiny,
            newAbe: simpleTemp.newAbe,
            size: simpleTemp.size,
          }));
        }}
      >
        {strings.cancel}
      </PrimaryButton>
    </div>
  );

  return (
    <div>
      <CommandButton
        onClick={() => {
          if (simple.itemID) {
            setSimpleTemp({
              urlImage: simple.urlImage,
              urlDestiny: simple.urlDestiny,
              newAbe: simple.newAbe,
              size: simple.size,
              alt: simple.alt,
            });
          }
          openPanel();
        }}
        text={`+ ${!simple.urlImage ? strings.newItem : strings.editItem}`}
        className={styles.buttonNewItem}
        styles={hostStyles}
      />
      <section className={styles.banner}>
        {simple.urlImage ? (
          <div className={styles.banner} onClick={onClickBanner}>
            <img
              className={
                simple.size === "25%"
                  ? styles.bannerImg25
                  : simple.size === "50%"
                  ? styles.bannerImg50
                  : simple.size === "75%"
                  ? styles.bannerImg75
                  : styles.bannerImg100
              }
              src={simple.urlImage}
              alt={simple.alt}
            />
          </div>
        ) : (
          <img
            className={styles.imagemExemploBanner}
            alt="Imagem sem Banner adicionado"
            src="https://whstorage2.blob.core.windows.net/brand/img/bannerPlace.svg"
          />
        )}
      </section>

      <Panel
        headerText={strings.appearance}
        headerClassName={styles.panelHeader}
        isOpen={isOpen}
        onDismiss={() => {
          dismissPanel();
          setPreview("");
        }}
        closeButtonAriaLabel={strings.close}
        isFooterAtBottom={true}
        onRenderFooterContent={onRenderFooterContent}
      >
        {preview && (
          <img
            className={styles.previewImg}
            src={preview}
            alt="Imagem a ser adicionada ao Banner"
          />
        )}
        <FilePicker
          buttonLabel={strings.chooseFile}
          label={strings.image}
          bingAPIKey="<BING API KEY>"
          accepts={[".gif", ".jpg", ".jpeg", ".png"]}
          buttonIcon="FileImage"
          onSave={(filePickerResult: IFilePickerResult[]) => {
            // eslint-disable-next-line @typescript-eslint/no-floating-promises
            changeImg(filePickerResult);
          }}
          context={props.context}
        />

        <Dropdown
          placeholder={simple.size ? simple.size : "100%"}
          label={strings.sizeOfImage}
          options={options}
          styles={dropdownStyles}
          onChange={changesize}
          defaultValue={simple.size}
        />

        <TextField
          type="text"
          onChange={changeurlDestiny}
          label={strings.urlDestiny}
          placeholder={simple.urlDestiny && simple.urlDestiny}
          value={simple.urlDestiny}
          styles={textFieldStyles}
        />

        <Toggle
          defaultChecked={simple.newAbe === false ? false : true}
          label={strings.openEmNewAbe}
          onText="Sim"
          offText="Não"
          onChange={changenewAbe}
          styles={toggleStyles}
        />

        <TextField
          type="text"
          onChange={changeAlt}
          label="Alt"
          placeholder={simple.alt && simple.alt}
          value={simple.alt}
          styles={textFieldStyles}
        />
      </Panel>
    </div>
  );
};

export default Simplebanner;
