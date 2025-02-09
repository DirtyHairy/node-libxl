var xl = require('../lib/libxl'),
    util = require('util'),
    testUtils = require('./testUtils'),
    shouldThrow = testUtils.shouldThrow;

describe('The CoreProperties class', () => {
    let book;
    let coreProperties;

    beforeEach(() => {
        book = new xl.Book(xl.BOOK_TYPE_XLSX);
        coreProperties = book.coreProperties();
    });

    it('title and setTitle change the title', () => {
        shouldThrow(coreProperties.title, {});
        shouldThrow(coreProperties.setTitle, coreProperties, 1);
        shouldThrow(coreProperties.setTitle, {}, 'Foo');

        expect(coreProperties.setTitle('Foo')).toBe(coreProperties);
        expect(coreProperties.title()).toBe('Foo');
    });

    it('subject and setSubject change the title', () => {
        shouldThrow(coreProperties.subject, {});
        shouldThrow(coreProperties.setSubject, coreProperties, 1);
        shouldThrow(coreProperties.setSubject, {}, 'Foo');

        expect(coreProperties.setSubject('Foo')).toBe(coreProperties);
        expect(coreProperties.subject()).toBe('Foo');
    });

    it('creator and setCreator change the title', () => {
        shouldThrow(coreProperties.creator, {});
        shouldThrow(coreProperties.setCreator, coreProperties, 1);
        shouldThrow(coreProperties.setCreator, {}, 'Foo');

        expect(coreProperties.setCreator('Foo')).toBe(coreProperties);
        expect(coreProperties.creator()).toBe('Foo');
    });

    it('lastModifiedBy and setLastModifiedBy change the title', () => {
        shouldThrow(coreProperties.lastModifiedBy, {});
        shouldThrow(coreProperties.setLastModifiedBy, coreProperties, 1);
        shouldThrow(coreProperties.setLastModifiedBy, {}, 'Foo');

        expect(coreProperties.setLastModifiedBy('Foo')).toBe(coreProperties);
        expect(coreProperties.lastModifiedBy()).toBe('Foo');
    });

    it('created and setCreated change the title', () => {
        const now = new Date().toISOString();

        shouldThrow(coreProperties.created, {});
        shouldThrow(coreProperties.setCreated, coreProperties, 1);
        shouldThrow(coreProperties.setCreated, {}, now);

        let date = new Date();
        expect(coreProperties.setCreated(now)).toBe(coreProperties);
        expect(coreProperties.created()).toBe(now);
    });

    it('createdAsDouble and setCreatedAsDouble change the title', () => {
        shouldThrow(coreProperties.createdAsDouble, {});
        shouldThrow(coreProperties.setCreatedAsDouble, coreProperties, 'a');
        shouldThrow(coreProperties.setCreatedAsDouble, {}, 1.0);

        expect(coreProperties.setCreatedAsDouble(1.0)).toBe(coreProperties);
        expect(coreProperties.createdAsDouble() - 1.0).toBeLessThan(testUtils.epsilon);
    });

    it('modified and setModified change the title', () => {
        const now = new Date().toISOString();

        shouldThrow(coreProperties.modified, {});
        shouldThrow(coreProperties.setModified, coreProperties, 1);
        shouldThrow(coreProperties.setModified, {}, now);

        let date = new Date();
        expect(coreProperties.setModified(now)).toBe(coreProperties);
        expect(coreProperties.modified()).toBe(now);
    });

    it('modifiedAsDouble and setModifiedAsDouble change the title', () => {
        shouldThrow(coreProperties.modifiedAsDouble, {});
        shouldThrow(coreProperties.setModifiedAsDouble, coreProperties, 'a');
        shouldThrow(coreProperties.setModifiedAsDouble, {}, 1.0);

        expect(coreProperties.setModifiedAsDouble(1.0)).toBe(coreProperties);
        expect(coreProperties.modifiedAsDouble() - 1.0).toBeLessThan(testUtils.epsilon);
    });

    it('tags and setTags change the title', () => {
        shouldThrow(coreProperties.tags, {});
        shouldThrow(coreProperties.setTags, coreProperties, 1);
        shouldThrow(coreProperties.setTags, {}, 'Foo');

        expect(coreProperties.setTags('Foo')).toBe(coreProperties);
        expect(coreProperties.tags()).toBe('Foo');
    });

    it('categories and setCategories change the title', () => {
        shouldThrow(coreProperties.categories, {});
        shouldThrow(coreProperties.setCategories, coreProperties, 1);
        shouldThrow(coreProperties.setCategories, {}, 'Foo');

        expect(coreProperties.setCategories('Foo')).toBe(coreProperties);
        expect(coreProperties.categories()).toBe('Foo');
    });

    it('comments and setComments change the title', () => {
        shouldThrow(coreProperties.comments, {});
        shouldThrow(coreProperties.setComments, coreProperties, 1);
        shouldThrow(coreProperties.setComments, {}, 'Foo');

        expect(coreProperties.setComments('Foo')).toBe(coreProperties);
        expect(coreProperties.comments()).toBe('Foo');
    });

    it('removeAll unsets all properties', () => {
        shouldThrow(coreProperties.removeAll, {});

        coreProperties.setTitle('Foo');

        expect(coreProperties.removeAll()).toBe(coreProperties);
        expect(() => coreProperties.title()).toThrow();
    });
});
