import { createFileRoute } from "@tanstack/react-router";
import { SF9MatatagPage } from "@/features/sf9-shared";

export const Route = createFileRoute("/_authenticated/sf9-matatag")({
  component: SF9MatatagPage,
});