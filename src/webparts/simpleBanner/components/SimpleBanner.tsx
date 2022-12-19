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
  fileName: string;
  urlImage: string;
  itemID: number;
  urlDestino: string;
  novaAba: boolean;
  tamanho: string;
}

interface SimpleBannerTemp {
  novaAba: boolean;
  urlDestino: string;
  tamanho: string;
}

const Simplebanner: React.FunctionComponent<ISimpleBannerProps> = (props) => {
  const sp = spfi().using(spSPFx(props.context));

  const simpleDefault: SimpleBanner = {
    fileName: null,
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
            fileName: res.FileLeafRef,
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

  const changeImg = async (file: IFilePickerResult[]): Promise<void> => {
    const oFile = file[0];
    if (oFile.fileAbsoluteUrl && oFile.fileAbsoluteUrl.length > 0) {
      setPreview(oFile.fileAbsoluteUrl);
    } else {
      setPreview(oFile.previewDataUrl);
    }
    // setSimple((old) => ({
    //   ...old,
    //   urlImage: preview,
    // }));
    setFile(oFile);
  };

  const Save = async (): Promise<void> => {
    if (preview) {
      const fileResultContent = await file.downloadFileContent();
      const result = await sp.web
        .getFolderByServerRelativePath("SimpleBanners")
        .files.addUsingPath(
          simple.fileName ? simple.fileName : fileResultContent.name,
          fileResultContent,
          {
            Overwrite: true,
          }
        );

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
          const newLink = res.FileRef.replace(
            res.FileLeafRef,
            fileResultContent.name
          );
          // eslint-disable-next-line @typescript-eslint/no-floating-promises
          sp.web.lists
            .getByTitle("SimpleBanners")
            .items.getById(res.Id)
            .update({
              FileLeafRef: fileResultContent.name,
              FileRef: newLink,
            })
            .then(() => {
              setSimple({
                fileName: fileResultContent.name,
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
              src={`${simple.urlImage}?p=${new Date().getTime()}`}
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
          label="Url do destino"
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
