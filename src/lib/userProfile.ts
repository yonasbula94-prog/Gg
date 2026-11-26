import { doc, getDoc, serverTimestamp, setDoc } from 'firebase/firestore';
import type { User } from 'firebase/auth';
import { db } from './firebase';

/**
 * Creates or refreshes the signed-in user's profile at users/{uid} in Firestore.
 * Matches the rules in firestore.rules: uid and email must match the auth token,
 * createdAt is set once, and updatedAt uses a server timestamp.
 */
export async function ensureUserProfile(user: User, displayName?: string): Promise<void> {
  const ref = doc(db, 'users', user.uid);
  const email = user.email ?? '';
  const name = (displayName || user.displayName || email.split('@')[0] || 'User').slice(0, 200);

  const snap = await getDoc(ref);
  if (snap.exists()) {
    await setDoc(
      ref,
      { uid: user.uid, email, displayName: name, updatedAt: serverTimestamp() },
      { merge: true }
    );
  } else {
    await setDoc(ref, {
      uid: user.uid,
      email,
      displayName: name,
      createdAt: serverTimestamp(),
      updatedAt: serverTimestamp(),
    });
  }
}
