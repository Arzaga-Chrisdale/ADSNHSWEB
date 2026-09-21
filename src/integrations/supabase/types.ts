export type Json =
  | string
  | number
  | boolean
  | null
  | { [key: string]: Json | undefined }
  | Json[];

export type Database = {
  // Allows to automatically instantiate createClient with right options
  // instead of createClient<Database, { PostgrestVersion: 'XX' }>(URL, KEY)
  __InternalSupabase: {
    PostgrestVersion: "14.5";
  };
  public: {
    Tables: {
      activity_scores: {
        Row: {
          activity_id: string;
          class_id: string;
          created_at: string;
          id: string;
          score: number | null;
          student_id: string;
          teacher_id: string;
        };
        Insert: {
          activity_id: string;
          class_id: string;
          created_at?: string;
          id?: string;
          score?: number | null;
          student_id: string;
          teacher_id: string;
        };
        Update: {
          activity_id?: string;
          class_id?: string;
          created_at?: string;
          id?: string;
          score?: number | null;
          student_id?: string;
          teacher_id?: string;
        };
        Relationships: [
          {
            foreignKeyName: "activity_scores_activity_id_fkey";
            columns: ["activity_id"];
            isOneToOne: false;
            referencedRelation: "grade_activities";
            referencedColumns: ["id"];
          },
          {
            foreignKeyName: "activity_scores_student_id_fkey";
            columns: ["student_id"];
            isOneToOne: false;
            referencedRelation: "students";
            referencedColumns: ["id"];
          },
        ];
      };
      attendance_records: {
        Row: {
          class_id: string;
          date: string;
          id: string;
          reason: string | null;
          status: string;
          student_id: string;
          teacher_id: string;
        };
        Insert: {
          class_id: string;
          date: string;
          id?: string;
          reason?: string | null;
          status: string;
          student_id: string;
          teacher_id: string;
        };
        Update: {
          class_id?: string;
          date?: string;
          id?: string;
          reason?: string | null;
          status?: string;
          student_id?: string;
          teacher_id?: string;
        };
        Relationships: [
          {
            foreignKeyName: "attendance_records_class_id_fkey";
            columns: ["class_id"];
            isOneToOne: false;
            referencedRelation: "classes";
            referencedColumns: ["id"];
          },
          {
            foreignKeyName: "attendance_records_student_id_fkey";
            columns: ["student_id"];
            isOneToOne: false;
            referencedRelation: "students";
            referencedColumns: ["id"];
          },
        ];
      };
      classes: {
        Row: {
          created_at: string;
          district: string | null;
          division: string | null;
          end_date: string | null;
          grade_level: string;
          id: string;
          region: string | null;
          school_id: string | null;
          school_name: string | null;
          school_year: string;
          section: string;
          start_date: string | null;
          subject: string;
          teacher_id: string;
          teacher_name: string | null;
          units: number | null;
          updated_at: string;
        };
        Insert: {
          created_at?: string;
          district?: string | null;
          division?: string | null;
          end_date?: string | null;
          grade_level: string;
          id?: string;
          region?: string | null;
          school_id?: string | null;
          school_name?: string | null;
          school_year: string;
          section: string;
          start_date?: string | null;
          subject: string;
          teacher_id: string;
          teacher_name?: string | null;
          units?: number | null;
          updated_at?: string;
        };
        Update: {
          created_at?: string;
          district?: string | null;
          division?: string | null;
          end_date?: string | null;
          grade_level?: string;
          id?: string;
          region?: string | null;
          school_id?: string | null;
          school_name?: string | null;
          school_year?: string;
          section?: string;
          start_date?: string | null;
          subject?: string;
          teacher_id?: string;
          teacher_name?: string | null;
          units?: number | null;
          updated_at?: string;
        };
        Relationships: [];
      };
      class_form_readiness: {
        Row: {
          adviser_id: string;
          class_id: string;
          form_code: string;
          is_ready: boolean;
          updated_at: string;
        };
        Insert: {
          adviser_id: string;
          class_id: string;
          form_code: string;
          is_ready?: boolean;
          updated_at?: string;
        };
        Update: {
          adviser_id?: string;
          class_id?: string;
          form_code?: string;
          is_ready?: boolean;
          updated_at?: string;
        };
        Relationships: [
          {
            foreignKeyName: "class_form_readiness_class_id_fkey";
            columns: ["class_id"];
            isOneToOne: false;
            referencedRelation: "classes";
            referencedColumns: ["id"];
          },
        ];
      };
      community_comments: {
        Row: {
          author_id: string;
          author_name: string | null;
          content: string;
          created_at: string;
          id: string;
          post_id: string;
        };
        Insert: {
          author_id: string;
          author_name?: string | null;
          content: string;
          created_at?: string;
          id?: string;
          post_id: string;
        };
        Update: {
          author_id?: string;
          author_name?: string | null;
          content?: string;
          created_at?: string;
          id?: string;
          post_id?: string;
        };
        Relationships: [
          {
            foreignKeyName: "community_comments_post_id_fkey";
            columns: ["post_id"];
            isOneToOne: false;
            referencedRelation: "community_posts";
            referencedColumns: ["id"];
          },
        ];
      };
      community_posts: {
        Row: {
          author_id: string;
          author_name: string | null;
          content: string;
          created_at: string;
          id: string;
          title: string;
        };
        Insert: {
          author_id: string;
          author_name?: string | null;
          content: string;
          created_at?: string;
          id?: string;
          title: string;
        };
        Update: {
          author_id?: string;
          author_name?: string | null;
          content?: string;
          created_at?: string;
          id?: string;
          title?: string;
        };
        Relationships: [];
      };
      grade_activities: {
        Row: {
          class_id: string;
          component: string;
          created_at: string;
          hps: number;
          id: string;
          position: number;
          teacher_id: string;
          term: string;
          title: string | null;
        };
        Insert: {
          class_id: string;
          component: string;
          created_at?: string;
          hps?: number;
          id?: string;
          position: number;
          teacher_id: string;
          term: string;
          title?: string | null;
        };
        Update: {
          class_id?: string;
          component?: string;
          created_at?: string;
          hps?: number;
          id?: string;
          position?: number;
          teacher_id?: string;
          term?: string;
          title?: string | null;
        };
        Relationships: [
          {
            foreignKeyName: "grade_activities_class_id_fkey";
            columns: ["class_id"];
            isOneToOne: false;
            referencedRelation: "classes";
            referencedColumns: ["id"];
          },
        ];
      };
      grade_components: {
        Row: {
          class_id: string;
          component: string;
          created_at: string;
          id: string;
          teacher_id: string;
          term: string;
          weight: number;
        };
        Insert: {
          class_id: string;
          component: string;
          created_at?: string;
          id?: string;
          teacher_id: string;
          term: string;
          weight?: number;
        };
        Update: {
          class_id?: string;
          component?: string;
          created_at?: string;
          id?: string;
          teacher_id?: string;
          term?: string;
          weight?: number;
        };
        Relationships: [
          {
            foreignKeyName: "grade_components_class_id_fkey";
            columns: ["class_id"];
            isOneToOne: false;
            referencedRelation: "classes";
            referencedColumns: ["id"];
          },
        ];
      };
      grade_requests: {
        Row: {
          class_id: string;
          created_at: string;
          id: string;
          status: string;
          subject: string;
          teacher_email: string;
          teacher_id: string;
          term: string;
        };
        Insert: {
          class_id: string;
          created_at?: string;
          id?: string;
          status?: string;
          subject: string;
          teacher_email: string;
          teacher_id: string;
          term: string;
        };
        Update: {
          class_id?: string;
          created_at?: string;
          id?: string;
          status?: string;
          subject?: string;
          teacher_email?: string;
          teacher_id?: string;
          term?: string;
        };
        Relationships: [
          {
            foreignKeyName: "grade_requests_class_id_fkey";
            columns: ["class_id"];
            isOneToOne: false;
            referencedRelation: "classes";
            referencedColumns: ["id"];
          },
        ];
      };
      grades: {
        Row: {
          class_id: string;
          id: string;
          score: number | null;
          student_id: string;
          subject: string;
          teacher_id: string;
          term: string;
        };
        Insert: {
          class_id: string;
          id?: string;
          score?: number | null;
          student_id: string;
          subject: string;
          teacher_id: string;
          term: string;
        };
        Update: {
          class_id?: string;
          id?: string;
          score?: number | null;
          student_id?: string;
          subject?: string;
          teacher_id?: string;
          term?: string;
        };
        Relationships: [
          {
            foreignKeyName: "grades_class_id_fkey";
            columns: ["class_id"];
            isOneToOne: false;
            referencedRelation: "classes";
            referencedColumns: ["id"];
          },
          {
            foreignKeyName: "grades_student_id_fkey";
            columns: ["student_id"];
            isOneToOne: false;
            referencedRelation: "students";
            referencedColumns: ["id"];
          },
        ];
      };
      learner_status_letters: {
        Row: {
          action_plan_items: Json;
          created_at: string;
          id: string;
          meeting_date: string | null;
          meeting_time: string | null;
          meeting_venue: string | null;
          student_id: string;
          teacher_id: string;
          teacher_notes: string | null;
        };
        Insert: {
          action_plan_items?: Json;
          created_at?: string;
          id?: string;
          meeting_date?: string | null;
          meeting_time?: string | null;
          meeting_venue?: string | null;
          student_id: string;
          teacher_id: string;
          teacher_notes?: string | null;
        };
        Update: {
          action_plan_items?: Json;
          created_at?: string;
          id?: string;
          meeting_date?: string | null;
          meeting_time?: string | null;
          meeting_venue?: string | null;
          student_id?: string;
          teacher_id?: string;
          teacher_notes?: string | null;
        };
        Relationships: [
          {
            foreignKeyName: "learner_status_letters_student_id_fkey";
            columns: ["student_id"];
            isOneToOne: false;
            referencedRelation: "students";
            referencedColumns: ["id"];
          },
        ];
      };
      profiles: {
        Row: {
          created_at: string;
          district: string | null;
          division: string | null;
          email: string | null;
          full_name: string | null;
          id: string;
          principal: string | null;
          region: string | null;
          school_id: string | null;
          school_name: string | null;
          teacher_type: string | null;
          updated_at: string;
        };
        Insert: {
          created_at?: string;
          district?: string | null;
          division?: string | null;
          email?: string | null;
          full_name?: string | null;
          id: string;
          principal?: string | null;
          region?: string | null;
          school_id?: string | null;
          school_name?: string | null;
          teacher_type?: string | null;
          updated_at?: string;
        };
        Update: {
          created_at?: string;
          district?: string | null;
          division?: string | null;
          email?: string | null;
          full_name?: string | null;
          id?: string;
          principal?: string | null;
          region?: string | null;
          school_id?: string | null;
          school_name?: string | null;
          teacher_type?: string | null;
          updated_at?: string;
        };
        Relationships: [];
      };
      school_form_submissions: {
        Row: {
          admin_remarks: string | null;
          adviser_id: string;
          adviser_name: string | null;
          class_id: string;
          grade_level: string;
          id: string;
          reviewed_at: string | null;
          reviewed_by: string | null;
          school_year: string | null;
          section: string | null;
          snapshot: Json;
          status: string;
          submitted_at: string;
        };
        Insert: {
          admin_remarks?: string | null;
          adviser_id: string;
          adviser_name?: string | null;
          class_id: string;
          grade_level: string;
          id?: string;
          reviewed_at?: string | null;
          reviewed_by?: string | null;
          school_year?: string | null;
          section?: string | null;
          snapshot?: Json;
          status?: string;
          submitted_at?: string;
        };
        Update: {
          admin_remarks?: string | null;
          adviser_id?: string;
          adviser_name?: string | null;
          class_id?: string;
          grade_level?: string;
          id?: string;
          reviewed_at?: string | null;
          reviewed_by?: string | null;
          school_year?: string | null;
          section?: string | null;
          snapshot?: Json;
          status?: string;
          submitted_at?: string;
        };
        Relationships: [
          {
            foreignKeyName: "school_form_submissions_class_id_fkey";
            columns: ["class_id"];
            isOneToOne: false;
            referencedRelation: "classes";
            referencedColumns: ["id"];
          },
        ];
      };
      students: {
        Row: {
          address: string | null;
          birthdate: string | null;
          class_id: string;
          contact_number: string | null;
          created_at: string;
          father_name: string | null;
          first_name: string;
          guardian: string | null;
          id: string;
          last_name: string;
          lrn: string | null;
          middle_name: string | null;
          mother_name: string | null;
          sex: string | null;
          teacher_id: string;
        };
        Insert: {
          address?: string | null;
          birthdate?: string | null;
          class_id: string;
          contact_number?: string | null;
          created_at?: string;
          father_name?: string | null;
          first_name: string;
          guardian?: string | null;
          id?: string;
          last_name: string;
          lrn?: string | null;
          middle_name?: string | null;
          mother_name?: string | null;
          sex?: string | null;
          teacher_id: string;
        };
        Update: {
          address?: string | null;
          birthdate?: string | null;
          class_id?: string;
          contact_number?: string | null;
          created_at?: string;
          father_name?: string | null;
          first_name?: string;
          guardian?: string | null;
          id?: string;
          last_name?: string;
          lrn?: string | null;
          middle_name?: string | null;
          mother_name?: string | null;
          sex?: string | null;
          teacher_id?: string;
        };
        Relationships: [
          {
            foreignKeyName: "students_class_id_fkey";
            columns: ["class_id"];
            isOneToOne: false;
            referencedRelation: "classes";
            referencedColumns: ["id"];
          },
        ];
      };
      user_roles: {
        Row: {
          created_at: string;
          id: string;
          role: string;
          user_id: string;
        };
        Insert: {
          created_at?: string;
          id?: string;
          role: string;
          user_id: string;
        };
        Update: {
          created_at?: string;
          id?: string;
          role?: string;
          user_id?: string;
        };
        Relationships: [];
      };
    };
    Views: {
      [_ in never]: never;
    };
    Functions: {
      admin_delete_user: {
        Args: { target_user_id: string };
        Returns: undefined;
      };
      is_admin: {
        Args: { check_user_id?: string };
        Returns: boolean;
      };
      review_school_form_submission: {
        Args: {
          p_admin_remarks?: string;
          p_status: string;
          p_submission_id: string;
        };
        Returns: undefined;
      };
      set_class_form_readiness: {
        Args: {
          p_class_id: string;
          p_form_code: string;
          p_is_ready: boolean;
        };
        Returns: undefined;
      };
      submit_school_forms_to_admin: {
        Args: { p_class_id: string };
        Returns: string;
      };
    };
    Enums: {
      [_ in never]: never;
    };
    CompositeTypes: {
      [_ in never]: never;
    };
  };
};

type DatabaseWithoutInternals = Omit<Database, "__InternalSupabase">;

type DefaultSchema = DatabaseWithoutInternals[Extract<
  keyof Database,
  "public"
>];

export type Tables<
  DefaultSchemaTableNameOrOptions extends
    | keyof (DefaultSchema["Tables"] & DefaultSchema["Views"])
    | { schema: keyof DatabaseWithoutInternals },
  TableName extends (DefaultSchemaTableNameOrOptions extends {
    schema: keyof DatabaseWithoutInternals;
  }
    ? keyof (DatabaseWithoutInternals[DefaultSchemaTableNameOrOptions["schema"]]["Tables"] &
        DatabaseWithoutInternals[DefaultSchemaTableNameOrOptions["schema"]]["Views"])
    : never) = never,
> = DefaultSchemaTableNameOrOptions extends {
  schema: keyof DatabaseWithoutInternals;
}
  ? (DatabaseWithoutInternals[DefaultSchemaTableNameOrOptions["schema"]]["Tables"] &
      DatabaseWithoutInternals[DefaultSchemaTableNameOrOptions["schema"]]["Views"])[TableName] extends {
      Row: infer R;
    }
    ? R
    : never
  : DefaultSchemaTableNameOrOptions extends keyof (DefaultSchema["Tables"] &
        DefaultSchema["Views"])
    ? (DefaultSchema["Tables"] &
        DefaultSchema["Views"])[DefaultSchemaTableNameOrOptions] extends {
        Row: infer R;
      }
      ? R
      : never
    : never;

export type TablesInsert<
  DefaultSchemaTableNameOrOptions extends
    keyof DefaultSchema["Tables"] | { schema: keyof DatabaseWithoutInternals },
  TableName extends (DefaultSchemaTableNameOrOptions extends {
    schema: keyof DatabaseWithoutInternals;
  }
    ? keyof DatabaseWithoutInternals[DefaultSchemaTableNameOrOptions["schema"]]["Tables"]
    : never) = never,
> = DefaultSchemaTableNameOrOptions extends {
  schema: keyof DatabaseWithoutInternals;
}
  ? DatabaseWithoutInternals[DefaultSchemaTableNameOrOptions["schema"]]["Tables"][TableName] extends {
      Insert: infer I;
    }
    ? I
    : never
  : DefaultSchemaTableNameOrOptions extends keyof DefaultSchema["Tables"]
    ? DefaultSchema["Tables"][DefaultSchemaTableNameOrOptions] extends {
        Insert: infer I;
      }
      ? I
      : never
    : never;

export type TablesUpdate<
  DefaultSchemaTableNameOrOptions extends
    keyof DefaultSchema["Tables"] | { schema: keyof DatabaseWithoutInternals },
  TableName extends (DefaultSchemaTableNameOrOptions extends {
    schema: keyof DatabaseWithoutInternals;
  }
    ? keyof DatabaseWithoutInternals[DefaultSchemaTableNameOrOptions["schema"]]["Tables"]
    : never) = never,
> = DefaultSchemaTableNameOrOptions extends {
  schema: keyof DatabaseWithoutInternals;
}
  ? DatabaseWithoutInternals[DefaultSchemaTableNameOrOptions["schema"]]["Tables"][TableName] extends {
      Update: infer U;
    }
    ? U
    : never
  : DefaultSchemaTableNameOrOptions extends keyof DefaultSchema["Tables"]
    ? DefaultSchema["Tables"][DefaultSchemaTableNameOrOptions] extends {
        Update: infer U;
      }
      ? U
      : never
    : never;

export type Enums<
  DefaultSchemaEnumNameOrOptions extends
    keyof DefaultSchema["Enums"] | { schema: keyof DatabaseWithoutInternals },
  EnumName extends (DefaultSchemaEnumNameOrOptions extends {
    schema: keyof DatabaseWithoutInternals;
  }
    ? keyof DatabaseWithoutInternals[DefaultSchemaEnumNameOrOptions["schema"]]["Enums"]
    : never) = never,
> = DefaultSchemaEnumNameOrOptions extends {
  schema: keyof DatabaseWithoutInternals;
}
  ? DatabaseWithoutInternals[DefaultSchemaEnumNameOrOptions["schema"]]["Enums"][EnumName]
  : DefaultSchemaEnumNameOrOptions extends keyof DefaultSchema["Enums"]
    ? DefaultSchema["Enums"][DefaultSchemaEnumNameOrOptions]
    : never;

export type CompositeTypes<
  PublicCompositeTypeNameOrOptions extends
    | keyof DefaultSchema["CompositeTypes"]
    | { schema: keyof DatabaseWithoutInternals },
  CompositeTypeName extends (PublicCompositeTypeNameOrOptions extends {
    schema: keyof DatabaseWithoutInternals;
  }
    ? keyof DatabaseWithoutInternals[PublicCompositeTypeNameOrOptions["schema"]]["CompositeTypes"]
    : never) = never,
> = PublicCompositeTypeNameOrOptions extends {
  schema: keyof DatabaseWithoutInternals;
}
  ? DatabaseWithoutInternals[PublicCompositeTypeNameOrOptions["schema"]]["CompositeTypes"][CompositeTypeName]
  : PublicCompositeTypeNameOrOptions extends keyof DefaultSchema["CompositeTypes"]
    ? DefaultSchema["CompositeTypes"][PublicCompositeTypeNameOrOptions]
    : never;

export const Constants = {
  public: {
    Enums: {},
  },
} as const;
