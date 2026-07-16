import { defineCollection, z } from 'astro:content';
import { glob } from 'astro/loaders';

const outfitters = defineCollection({
  loader: glob({ pattern: '**/*.md', base: './src/content/outfitters' }),
  schema: z.object({
    name: z.string(),
    baseTown: z.string(),
    region: z.string(),
    waters: z.array(z.string()),
    fishingType: z.array(z.string()),
    priceTier: z.enum(['$', '$$', '$$$', 'Contact for pricing']),
    priceDisplay: z.string(),
    phone: z.string().optional(),
    website: z.string().optional(),
    tripTypes: z.string().optional(),
    speciesTargeted: z.string().optional(),
    seasonActive: z.string().optional(),
    description: z.string(),
    photo: z.string().optional(),
    verified: z.boolean(),
    leadStatus: z.string(),
  }),
});

const properties = defineCollection({
  loader: glob({ pattern: '**/*.md', base: './src/content/properties' }),
  schema: z.object({
    name: z.string(),
    town: z.string(),
    region: z.string(),
    propertyType: z.string(),
    priceTier: z.enum(['$', '$$', '$$$', '$$$$']),
    priceDisplay: z.string(),
    amenities: z.array(z.string()),
    roomTypes: z.string().optional(),
    phone: z.string().optional(),
    website: z.string().optional(),
    affiliateUrl: z.string().optional(),
    description: z.string(),
    photo: z.string().optional(),
    verified: z.boolean(),
  }),
});

export const collections = {
  outfitters,
  properties,
};