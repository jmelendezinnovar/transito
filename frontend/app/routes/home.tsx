import type { Route } from "./+types/home";
import { Welcome } from "../welcome/welcome";

export function meta({}: Route.MetaArgs) {
  return [
    { title: "Flujo de Transito Cartera" },
    { name: "description", content: "Visualización del flujo de procesamiento de archivos" },
  ];
}

export default function Home() {
  return <Welcome />;
}
