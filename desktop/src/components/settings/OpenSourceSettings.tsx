/**
 * Open source project usage details.
 *
 * Lists direct runtime and build dependencies used by this desktop app.
 */

import React from "react";
import { ExternalLink, Package } from "../common/IconMap";
import { useSettingsStore } from "../../store/settingsStore";
import { ipcApi } from "../../services/ipcApi";
import { OPEN_SOURCE_PROJECTS, type OpenSourceProject } from "./openSourceSettingsContent";

const OPEN_SOURCE_TEXT = {
  "zh-CN": {
    title: "开源项目",
    desc: "本页用于声明软件中直接使用的主要开源项目、用途和项目地址。",
    noticeTitle: "开源使用声明",
    noticeBody:
      "以下项目为本应用的直接运行依赖、构建工具、外部解析服务或开放格式参考。各项目版权和许可证归原作者及社区所有，具体许可条款以对应仓库、官方标准和随包 LICENSE 文件为准。",
    runtime: "运行时依赖",
    service: "外部服务与开放标准",
    build: "构建与测试工具",
    project: "项目",
    purpose: "主要用途",
    license: "许可证",
    link: "GitHub",
    open: "打开 GitHub",
  },
  "en-US": {
    title: "Open Source",
    desc: "This page declares the main open source projects directly used by the app, their purpose, and project links.",
    noticeTitle: "Open Source Notice",
    noticeBody:
      "The projects below are direct runtime dependencies, build tools, external parsing services, or open format references used by this app. Copyrights and licenses belong to their authors and communities; the authoritative terms are the linked repositories, official standards, and bundled LICENSE files.",
    runtime: "Runtime Dependencies",
    service: "External Services And Open Standards",
    build: "Build And Test Tools",
    project: "Project",
    purpose: "Purpose",
    license: "License",
    link: "GitHub",
    open: "Open GitHub",
  },
} as const;

export const OpenSourceSettings: React.FC = () => {
  const { language } = useSettingsStore();
  const text = OPEN_SOURCE_TEXT[language];
  const runtimeProjects = OPEN_SOURCE_PROJECTS.filter((project) => project.category === "runtime");
  const serviceProjects = OPEN_SOURCE_PROJECTS.filter(
    (project) => project.category === "service" || project.category === "standard",
  );
  const buildProjects = OPEN_SOURCE_PROJECTS.filter((project) => project.category === "build");

  const openGithub = async (url: string) => {
    await ipcApi.app.openExternal(url);
  };

  return (
    <div className="settings-section-content open-source-settings">
      <h2>{text.title}</h2>
      <p className="section-desc">{text.desc}</p>

      <div className="settings-card open-source-notice">
        <div className="settings-card-header">
          <div className="settings-card-title-row">
            <Package size={16} />
            <h3>{text.noticeTitle}</h3>
          </div>
          <p>{text.noticeBody}</p>
        </div>
      </div>

      <OpenSourceGroup
        title={text.runtime}
        projects={runtimeProjects}
        text={text}
        onOpenGithub={openGithub}
        language={language}
      />
      <OpenSourceGroup
        title={text.service}
        projects={serviceProjects}
        text={text}
        onOpenGithub={openGithub}
        language={language}
      />
      <OpenSourceGroup
        title={text.build}
        projects={buildProjects}
        text={text}
        onOpenGithub={openGithub}
        language={language}
      />
    </div>
  );
};

function OpenSourceGroup({
  title,
  projects,
  text,
  onOpenGithub,
  language,
}: {
  title: string;
  projects: OpenSourceProject[];
  text: (typeof OPEN_SOURCE_TEXT)["zh-CN"] | (typeof OPEN_SOURCE_TEXT)["en-US"];
  onOpenGithub: (url: string) => void;
  language: "zh-CN" | "en-US";
}) {
  return (
    <div className="settings-card open-source-group">
      <div className="settings-card-header">
        <h3>{title}</h3>
      </div>
      <div className="open-source-table-wrapper">
        <table className="open-source-table">
          <thead>
            <tr>
              <th>{text.project}</th>
              <th>{text.purpose}</th>
              <th>{text.license}</th>
              <th>{text.link}</th>
            </tr>
          </thead>
          <tbody>
            {projects.map((project) => (
              <tr key={project.name}>
                <td className="open-source-name">{project.name}</td>
                <td>{language === "zh-CN" ? project.purposeZh : project.purposeEn}</td>
                <td>
                  <span className="open-source-license">{project.license}</span>
                </td>
                <td>
                  <button
                    type="button"
                    className="open-source-link"
                    onClick={() => onOpenGithub(project.githubUrl)}
                    title={text.open}
                  >
                    <ExternalLink size={14} />
                    <span>{text.link}</span>
                  </button>
                </td>
              </tr>
            ))}
          </tbody>
        </table>
      </div>
    </div>
  );
}
