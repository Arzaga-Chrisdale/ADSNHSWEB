import { createFileRoute } from "@tanstack/react-router";
import { useQuery, useMutation, useQueryClient } from "@tanstack/react-query";
import { supabase } from "@/integrations/supabase/client";
import { Button } from "@/components/ui/button";
import { Input } from "@/components/ui/input";
import { Label } from "@/components/ui/label";
import { useEffect, useState } from "react";
import { toast } from "sonner";
import { getUserId } from "@/lib/data";

export const Route = createFileRoute("/_authenticated/school-info")({
  component: SchoolInfo,
});

function SchoolInfo() {
  const qc = useQueryClient();
  const { data: profile } = useQuery({ queryKey: ["profile"], queryFn: async () => (await supabase.from("profiles").select("*").maybeSingle()).data });
  const [f, setF] = useState({ full_name: "", school_name: "", school_id: "", region: "", division: "", district: "", principal: "" });
  useEffect(() => { if (profile) setF({
    full_name: profile.full_name ?? "", school_name: profile.school_name ?? "", school_id: profile.school_id ?? "",
    region: profile.region ?? "", division: profile.division ?? "", district: profile.district ?? "", principal: profile.principal ?? "",
  }); }, [profile]);

  const save = useMutation({
    mutationFn: async () => {
      const id = await getUserId();
      const { error } = await supabase.from("profiles").upsert({ id, ...f, updated_at: new Date().toISOString() });
      if (error) throw error;
    },
    onSuccess: () => { toast.success("Saved"); qc.invalidateQueries({ queryKey: ["profile"] }); },
    onError: (e: Error) => toast.error(e.message),
  });

  return (
    <div className="space-y-4 pb-24 md:pb-4">
      <div>
        <h1 className="text-2xl font-semibold">School Info</h1>
        <p className="text-sm text-muted-foreground">Fills the header of every DepEd form and letter.</p>
      </div>
      <div className="rounded-2xl border bg-card p-4 grid gap-3 sm:grid-cols-2">
        <div><Label>Your name</Label><Input value={f.full_name} onChange={(e) => setF({ ...f, full_name: e.target.value })} /></div>
        <div><Label>Principal</Label><Input value={f.principal} onChange={(e) => setF({ ...f, principal: e.target.value })} /></div>
        <div><Label>School name</Label><Input value={f.school_name} onChange={(e) => setF({ ...f, school_name: e.target.value })} /></div>
        <div><Label>School ID</Label><Input value={f.school_id} onChange={(e) => setF({ ...f, school_id: e.target.value })} /></div>
        <div><Label>Region</Label><Input value={f.region} onChange={(e) => setF({ ...f, region: e.target.value })} /></div>
        <div><Label>Division</Label><Input value={f.division} onChange={(e) => setF({ ...f, division: e.target.value })} /></div>
        <div className="sm:col-span-2"><Label>District</Label><Input value={f.district} onChange={(e) => setF({ ...f, district: e.target.value })} /></div>
        <div className="sm:col-span-2 flex justify-end"><Button onClick={() => save.mutate()}>Save</Button></div>
      </div>
    </div>
  );
}
