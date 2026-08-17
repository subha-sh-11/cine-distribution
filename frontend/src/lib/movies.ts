"use client";

import { useSyncExternalStore } from "react";

export type Movie = {
  id: string;
  name: string;
  release: string; // free text / date
  createdAt: number;
  owned?: boolean; // false → shared with me (not my upload)
  ownerEmail?: string | null; // who uploaded/owns it
  sharedBy?: string | null; // who actually shared it with me (never the admin acct)
  role?: string; // my access: "editor" | "viewer"
};

// Movies are stored in the shared Postgres DB (via /api/movies) so every admin
// user sees the same list. We keep a small in-memory cache + subscription so the
// UI updates reactively, and refetch from the server to stay in sync.

let cache: Movie[] = [];
let loaded = false;
const listeners = new Set<() => void>();

function emit() {
  listeners.forEach((l) => l());
}

export async function refreshMovies(): Promise<Movie[]> {
  try {
    const res = await fetch("/api/movies", { cache: "no-store" });
    if (res.ok) {
      cache = (await res.json()).movies ?? [];
      loaded = true;
      emit();
    }
  } catch {
    /* keep last cache */
  }
  return cache;
}

function subscribe(cb: () => void) {
  listeners.add(cb);
  if (!loaded) refreshMovies(); // lazy first load
  return () => listeners.delete(cb);
}

/** Reactive shared movies list (from the DB). Empty until the first fetch. */
export function useMovies(): Movie[] {
  return useSyncExternalStore(
    subscribe,
    () => cache,
    () => cache
  );
}

export async function addMovie(name: string, release: string): Promise<string> {
  const res = await fetch("/api/movies", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ name, release }),
  });
  const data = await res.json().catch(() => ({}));
  const movie: Movie | undefined = data.movie;
  if (movie) {
    cache = [movie, ...cache];
    emit();
    return movie.id;
  }
  await refreshMovies();
  return "";
}

export async function deleteMovie(id: string): Promise<void> {
  cache = cache.filter((m) => m.id !== id); // optimistic
  emit();
  try {
    await fetch(`/api/movies/${id}`, { method: "DELETE" });
  } finally {
    refreshMovies();
  }
}

/** Look up a movie already in the cache (call refreshMovies first if unsure). */
export function getMovie(id: string): Movie | undefined {
  return cache.find((m) => m.id === id);
}
