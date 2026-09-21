import { createFileRoute, Navigate } from "@tanstack/react-router";

export const Route = createFileRoute("/")({
  component: LandingRedirect,
});

function LandingRedirect() {
  return <Navigate to="/auth" replace />;
}
