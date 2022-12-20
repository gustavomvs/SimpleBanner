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
  PanelType,
} from "@fluentui/react";

const textFieldStyles: Partial<ITextFieldStyles> = {
  fieldGroup: { width: "90%", marginBottom: "5px" },
};

const hostStyles: Partial<ITooltipHostStyles> = {
  root: { height: "20px", margin: "8px" },
};

const dropdownStyles: Partial<IDropdownStyles> = {
  dropdown: { width: "90%" },
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
  urlDestino: string;
  novaAba: boolean;
  tamanho: string;
}

interface SimpleBannerTemp {
  urlImage: string;
  novaAba: boolean;
  urlDestino: string;
  tamanho: string;
}

const Simplebanner: React.FunctionComponent<ISimpleBannerProps> = (props) => {
  const sp = spfi().using(spSPFx(props.context));

  const simpleDefault: SimpleBanner = {
    urlImage: null,
    itemID: props.itemId,
    urlDestino: "",
    novaAba: true,
    tamanho: "100%",
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
          "novaAba",
          "tamanho",
          "urlDestino",
          "FileLeafRef"
        )()
        .then((res) => {
          setSimple({
            urlImage: res.FileRef,
            itemID: res.Id,
            urlDestino: res.urlDestino,
            novaAba: res.novaAba,
            tamanho: res.tamanho,
          });
        });
    }
  };

  useEffect(() => {
    // eslint-disable-next-line @typescript-eslint/no-floating-promises
    meuInit();
  }, []);

  const changeTamanho = (
    event: React.FormEvent<HTMLDivElement>,
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    option?: any,
    index?: number
  ): void => {
    setSimple((old) => ({
      ...old,
      tamanho: option.text,
    }));
  };

  const changeUrlDestino = (
    event: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>,
    newValue?: string
  ): void => {
    setSimple((old) => ({
      ...old,
      urlDestino: newValue,
    }));
  };

  function changeNovaAba(
    _ev: React.MouseEvent<HTMLElement>,
    checked?: boolean
  ): void {
    setSimple((old) => ({
      ...old,
      novaAba: checked,
    }));
  }

  const changeImg = (file: IFilePickerResult[]): void => {
    const oFile = file[0];
    if (oFile.fileAbsoluteUrl && oFile.fileAbsoluteUrl.length > 0) {
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
        novaAba: simple.novaAba,
        urlDestino: simple.urlDestino,
        tamanho: simple.tamanho,
      });

      await item
        .select(
          "Id",
          "FileRef",
          "novaAba",
          "tamanho",
          "urlDestino",
          "FileLeafRef"
        )()
        .then((res) => {
          const nameSplit = res.FileLeafRef.split(".");
          const nameFinal = nameSplit[nameSplit.length - 1];
          const nameInitial = res.FileLeafRef.slice(
            0,
            res.FileLeafRef.length - nameFinal.length - 1
          ).slice(0, 200);
          const newLink = res.FileRef.replace(
            res.FileLeafRef,
            `${nameInitial}_${res.Id}_.${nameFinal}`
          );
          // eslint-disable-next-line @typescript-eslint/no-floating-promises
          sp.web.lists
            .getByTitle("SimpleBanners")
            .items.getById(res.Id)
            .update({
              FileLeafRef: `${nameInitial}_${res.Id}_`,
            })
            .then(() => {
              setSimple({
                urlImage: newLink,
                itemID: res.Id,
                urlDestino: res.urlDestino,
                novaAba: res.novaAba,
                tamanho: res.tamanho,
              });
            });

          props.updatePropety(res.Id);
        });
    } else {
      if (simple.itemID) {
        await sp.web.lists
          .getByTitle("SimpleBanners")
          .items.getById(simple.itemID)
          .update({
            novaAba: simple.novaAba === false ? false : true,
            urlDestino: simple.urlDestino,
            tamanho: simple.tamanho,
          })
          .then((res) => {
            setSimple((old) => ({
              ...old,
              urlDestino: simple.urlDestino,
              novaAba: simple.novaAba,
              tamanho: simple.tamanho,
            }));
          });
      }
    }

    dismissPanel();
    setPreview("");
  };

  const onClickBanner = (): void => {
    if (simple.urlDestino) {
      window.open(simple.urlDestino, simple.novaAba ? "_blank" : "_self");
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
        Salvar
      </DefaultButton>
      <PrimaryButton
        onClick={() => {
          dismissPanel();
          setPreview("");
          setSimple((old) => ({
            ...old,
            urlImage: simpleTemp.urlImage,
            urlDestino: simpleTemp.urlDestino,
            novaAba: simpleTemp.novaAba,
            tamanho: simpleTemp.tamanho,
          }));
        }}
      >
        Cancelar
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
              urlDestino: simple.urlDestino,
              novaAba: simple.novaAba,
              tamanho: simple.tamanho,
            });
          }
          openPanel();
        }}
        text={"+ Novo item"}
        className={styles.buttonNewItem}
        styles={hostStyles}
      />
      <section className={styles.banner}>
        {simple.urlImage ? (
          <div className={styles.banner} onClick={onClickBanner}>
            <img
              className={
                simple.tamanho === "25%"
                  ? styles.bannerImg25
                  : simple.tamanho === "50%"
                  ? styles.bannerImg50
                  : simple.tamanho === "75%"
                  ? styles.bannerImg75
                  : styles.bannerImg100
              }
              src={`${simple.urlImage}`}
              alt="Imagem"
            />
          </div>
        ) : (
          <img
            className={styles.imagemExemploBanner}
            alt="imagem sem nada"
            src="https://whstorage2.blob.core.windows.net/brand/img/bannerPlace.svg"
          />
        )}
      </section>

      <Panel
        type={PanelType.custom}
        customWidth={"500px"}
        headerText="Aparência"
        isOpen={isOpen}
        onDismiss={() => {
          dismissPanel();
          setPreview("");
        }}
        closeButtonAriaLabel="Close"
        isFooterAtBottom={true}
        onRenderFooterContent={onRenderFooterContent}
      >
        {preview && (
          <img
            className={styles.previewImg}
            src={preview}
            alt="Imagem a ser adicionada"
          />
        )}
        <FilePicker
          label="Imagem"
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
          placeholder={simple.tamanho ? simple.tamanho : "100%"}
          label="Tamanho da imagem"
          options={options}
          styles={dropdownStyles}
          onChange={changeTamanho}
          defaultValue={simple.tamanho}
        />

        <TextField
          type="text"
          onChange={changeUrlDestino}
          label="URL do destino"
          placeholder={simple.urlDestino && simple.urlDestino}
          value={simple.urlDestino}
          styles={textFieldStyles}
        />

        <Toggle
          defaultChecked={simple.novaAba === false ? false : true}
          label="Abrir em nova aba?"
          onText="Sim"
          offText="Não"
          onChange={changeNovaAba}
        />
      </Panel>
    </div>
  );
};

export default Simplebanner;
