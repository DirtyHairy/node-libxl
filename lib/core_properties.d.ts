export class CoreProperties {
    title(): string;
    setTitle(title: string): CoreProperties;
    subject(): string;
    setSubject(subject: string): CoreProperties;
    creator(): string;
    setCreator(creator: string): CoreProperties;
    lastModifiedBy(): string;
    setLastModifiedBy(lastModifiedBy: string): CoreProperties;
    created(): string;
    createdAsDouble(): number;
    setCreated(created: string): CoreProperties;
    setCreatedAsDouble(created: number): CoreProperties;
    modified(): string;
    modifiedAsDouble(): number;
    setModified(modified: string): CoreProperties;
    setModifiedAsDouble(modified: number): CoreProperties;
    tags(): string;
    setTags(tags: string): CoreProperties;
    categories(): string;
    setCategories(categories: string): CoreProperties;
    comments(): string;
    setComments(comments: string): CoreProperties;
    removeAll(): CoreProperties;
}
